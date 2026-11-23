using System.Collections.Concurrent;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Channels;
using ExcelDataReader;
using ExcelPagedSearchApp.Models;

namespace ExcelPagedSearchApp.Services;

public sealed class ExcelPagedSearcher
{
    static ExcelPagedSearcher()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    private static ExcelReaderConfiguration CreateConfig(string? password) =>
        new() { Password = password };

    public static IReadOnlyList<string> GetSheetNames(string filePath, string? password = null)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream, CreateConfig(password));

        var names = new List<string>();
        do
        {
            if (!string.IsNullOrWhiteSpace(reader.Name))
                names.Add(reader.Name);
        } while (reader.NextResult());

        return names;
    }

    public static IReadOnlyList<ColumnOption> GetColumnHeaders(
        string filePath,
        string? sheetName,
        bool hasHeaderRow = true,
        string? password = null)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream, CreateConfig(password));

        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            while (!string.Equals(reader.Name, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                if (!reader.NextResult())
                    return Array.Empty<ColumnOption>();
            }
        }

        if (!reader.Read())
            return Array.Empty<ColumnOption>();

        var columns = new List<ColumnOption>(reader.FieldCount);
        for (int col = 0; col < reader.FieldCount; col++)
        {
            string rawHeader = hasHeaderRow
                ? (reader.GetValue(col)?.ToString()?.Trim() ?? string.Empty)
                : string.Empty;

            string label = string.IsNullOrWhiteSpace(rawHeader)
                ? $"Column {col + 1}"
                : rawHeader;

            columns.Add(new ColumnOption { Index = col, HeaderName = label, IsSelected = true });
        }

        return columns;
    }

    public async IAsyncEnumerable<PageSearchResult> SearchByPagesAsync(
        string filePath,
        string? sheetName,
        string query,
        PageSearchOptions options,
        string? password = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path is required.", nameof(filePath));
        if (string.IsNullOrWhiteSpace(query))
            throw new ArgumentException("Search query is required.", nameof(query));
        if (options.PageSize <= 0)
            throw new ArgumentOutOfRangeException(nameof(options.PageSize), "Page size must be greater than zero.");

        var comparison = options.CaseSensitive
            ? StringComparison.Ordinal
            : StringComparison.OrdinalIgnoreCase;

        await using var stream = new FileStream(
            filePath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.ReadWrite,
            bufferSize: 65536,
            useAsync: true);

        using var reader = ExcelReaderFactory.CreateReader(stream, CreateConfig(password));

        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            while (!string.Equals(reader.Name, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                if (!reader.NextResult())
                    yield break;
            }
        }

        string[] headers = Array.Empty<string>();
        int dataRowIndex = 0;
        int pageNumber = 0;
        var currentPageMatches = new List<EditableRowMatch>();

        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();

            int currentRowNumber = reader.Depth + 1;

            if (options.HasHeaderRow && currentRowNumber == 1)
            {
                headers = new string[reader.FieldCount];
                for (int col = 0; col < reader.FieldCount; col++)
                {
                    headers[col] = reader.GetValue(col)?.ToString()?.Trim() ?? $"Column {col + 1}";
                }
                continue;
            }

            int targetPage = (dataRowIndex / options.PageSize) + 1;
            if (pageNumber == 0)
                pageNumber = targetPage;

            if (targetPage != pageNumber)
            {
                if (currentPageMatches.Count > 0)
                {
                    yield return new PageSearchResult(pageNumber, currentPageMatches.ToList());
                    currentPageMatches.Clear();
                }
                pageNumber = targetPage;
            }

            var match = TryBuildRowMatch(reader, headers, pageNumber, currentRowNumber, query, comparison, options);
            if (match is not null)
            {
                currentPageMatches.Add(match);

                if (options.StopOnFirstMatch)
                {
                    yield return new PageSearchResult(pageNumber, currentPageMatches.ToList());
                    yield break;
                }
            }

            dataRowIndex++;

            if (dataRowIndex % (options.PageSize * 25) == 0 && currentPageMatches.Count > 0)
            {
                yield return new PageSearchResult(pageNumber, currentPageMatches.ToList());
                currentPageMatches.Clear();
                await Task.Yield();
            }
        }

        if (currentPageMatches.Count > 0)
        {
            yield return new PageSearchResult(pageNumber == 0 ? 1 : pageNumber, currentPageMatches);
        }
    }

    public async IAsyncEnumerable<PageSearchResult> SearchDistributedSynchronizedAsync(
        string filePath,
        string? sheetName,
        string query,
        PageSearchOptions options,
        int pageStride = 4,
        string? password = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        if (pageStride <= 1)
        {
            await foreach (var res in SearchByPagesAsync(filePath, sheetName, query, options, password, cancellationToken))
                yield return res;
            yield break;
        }

        var channel = Channel.CreateUnbounded<PageSearchResult>(new UnboundedChannelOptions
        {
            SingleReader = true,
            SingleWriter = true
        });

        var producerTask = Task.Run(async () =>
        {
            var bucketedMatches = new ConcurrentDictionary<int, List<EditableRowMatch>>();

            try
            {
                await using var stream = new FileStream(
                    filePath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.ReadWrite,
                    bufferSize: 65536,
                    useAsync: true);

                using var reader = ExcelReaderFactory.CreateReader(stream, CreateConfig(password));

                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    while (!string.Equals(reader.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (!reader.NextResult())
                        {
                            channel.Writer.Complete();
                            return;
                        }
                    }
                }

                var comparison = options.CaseSensitive
                    ? StringComparison.Ordinal
                    : StringComparison.OrdinalIgnoreCase;

                string[] headers = Array.Empty<string>();
                int dataRowIndex = 0;
                int highestPageSeen = 0;
                int nextFlushPhase = 0;

                while (reader.Read())
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    int currentRowNumber = reader.Depth + 1;
                    if (options.HasHeaderRow && currentRowNumber == 1)
                    {
                        headers = new string[reader.FieldCount];
                        for (int col = 0; col < reader.FieldCount; col++)
                            headers[col] = reader.GetValue(col)?.ToString()?.Trim() ?? $"Column {col + 1}";
                        continue;
                    }

                    int pageNumber = (dataRowIndex / options.PageSize) + 1;
                    if (pageNumber > highestPageSeen)
                        highestPageSeen = pageNumber;

                    var match = TryBuildRowMatch(reader, headers, pageNumber, currentRowNumber, query, comparison, options);
                    if (match is not null)
                    {
                        var list = bucketedMatches.GetOrAdd(pageNumber, _ => new List<EditableRowMatch>());
                        list.Add(match);

                        if (options.StopOnFirstMatch)
                        {
                            await channel.Writer.WriteAsync(new PageSearchResult(pageNumber, new[] { match }), cancellationToken);
                            channel.Writer.Complete();
                            return;
                        }
                    }

                    dataRowIndex++;

                    if (dataRowIndex % (options.PageSize * pageStride) == 0)
                    {
                        await FlushDistributedPhaseAsync(
                            bucketedMatches,
                            highestPageSeen,
                            pageStride,
                            nextFlushPhase,
                            channel.Writer,
                            cancellationToken);

                        nextFlushPhase = (nextFlushPhase + 1) % pageStride;
                    }
                }

                for (int phase = 0; phase < pageStride; phase++)
                {
                    await FlushDistributedPhaseAsync(
                        bucketedMatches,
                        highestPageSeen,
                        pageStride,
                        phase,
                        channel.Writer,
                        cancellationToken);
                }

                channel.Writer.Complete();
            }
            catch (OperationCanceledException)
            {
                channel.Writer.TryComplete();
            }
            catch (Exception ex)
            {
                channel.Writer.TryComplete(ex);
            }
        }, cancellationToken);

        await foreach (var item in channel.Reader.ReadAllAsync(cancellationToken))
        {
            yield return item;
        }

        await producerTask;
    }

    private static EditableRowMatch? TryBuildRowMatch(
        IExcelDataReader reader,
        string[] headers,
        int pageNumber,
        int currentRowNumber,
        string query,
        StringComparison comparison,
        PageSearchOptions options)
    {
        bool filterSearchCols = options.SearchColumnIndices is { Count: > 0 };
        bool filterDisplayCols = options.DisplayColumnIndices is { Count: > 0 };

        var allCellTexts = new string[reader.FieldCount];
        var uniqueKeys = new string[reader.FieldCount];
        var seenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (int col = 0; col < reader.FieldCount; col++)
        {
            string baseHeader = col < headers.Length && !string.IsNullOrWhiteSpace(headers[col])
                ? headers[col]
                : $"Column {col + 1}";

            string key = seenKeys.Add(baseHeader)
                ? baseHeader
                : $"{baseHeader} ({col + 1})";

            uniqueKeys[col] = key;
            allCellTexts[col] = reader.GetValue(col)?.ToString() ?? string.Empty;
        }

        var matchedColumns = new List<string>();
        bool rowMatched = false;

        for (int col = 0; col < reader.FieldCount; col++)
        {
            if (filterSearchCols && !options.SearchColumnIndices!.Contains(col))
                continue;

            string cellText = allCellTexts[col];
            if (!string.IsNullOrEmpty(cellText) && cellText.Contains(query, comparison))
            {
                rowMatched = true;
                matchedColumns.Add(uniqueKeys[col]);
            }
        }

        if (!rowMatched)
            return null;

        var displayRowValues = new Dictionary<string, string>();
        var columnIndexMap = new Dictionary<string, int>();

        for (int col = 0; col < reader.FieldCount; col++)
        {
            if (filterDisplayCols && !options.DisplayColumnIndices!.Contains(col))
                continue;

            string key = uniqueKeys[col];
            displayRowValues[key] = allCellTexts[col];
            columnIndexMap[key] = col;
        }

        return new EditableRowMatch(
            pageNumber,
            currentRowNumber,
            string.Join(", ", matchedColumns),
            displayRowValues,
            columnIndexMap);
    }

    private static async Task FlushDistributedPhaseAsync(
        ConcurrentDictionary<int, List<EditableRowMatch>> bucketedMatches,
        int highestPageSeen,
        int pageStride,
        int phase,
        ChannelWriter<PageSearchResult> writer,
        CancellationToken ct)
    {
        for (int p = 1; p <= highestPageSeen; p++)
        {
            if ((p - 1) % pageStride != phase)
                continue;

            if (bucketedMatches.TryRemove(p, out var matches) && matches.Count > 0)
            {
                await writer.WriteAsync(new PageSearchResult(p, matches), ct);
            }
        }
    }
}