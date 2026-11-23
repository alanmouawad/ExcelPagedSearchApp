using System.IO;
using System.Text;
using ClosedXML.Excel;
using ExcelDataReader;
using ExcelPagedSearchApp.Models;

namespace ExcelPagedSearchApp.Services;

public sealed record SaveChangesResult(int UpdatedCells, string SavedFilePath, bool SavedAsNewCopy);

public sealed class ExcelWriterService
{
    static ExcelWriterService()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    public SaveChangesResult SaveModifiedRows(
        string filePath,
        string? sheetName,
        IEnumerable<EditableRowMatch> rows,
        string? password = null)
    {
        var modifiedRows = rows.Where(r => r.IsModified).ToList();
        if (modifiedRows.Count == 0)
            return new SaveChangesResult(0, filePath, SavedAsNewCopy: false);

        if (string.IsNullOrEmpty(password))
        {
            try
            {
                int inPlaceCount = SaveInPlace(filePath, sheetName, modifiedRows);
                return new SaveChangesResult(inPlaceCount, filePath, SavedAsNewCopy: false);
            }
            catch (Exception ex) when (IsEncryptionOrFormatLimitation(ex))
            {
                // Fall through to decrypted copy export if the workbook is encrypted
            }
        }

        string copyPath = BuildEditedCopyPath(filePath);
        int copyCount = SaveAsDecryptedCopy(filePath, copyPath, sheetName, modifiedRows, password);
        return new SaveChangesResult(copyCount, copyPath, SavedAsNewCopy: true);
    }

    private static int SaveInPlace(string filePath, string? sheetName, List<EditableRowMatch> modifiedRows)
    {
        using var workbook = new XLWorkbook(filePath);
        var worksheet = string.IsNullOrWhiteSpace(sheetName)
            ? workbook.Worksheets.First()
            : workbook.Worksheet(sheetName);

        int updatedCells = 0;

        foreach (var row in modifiedRows)
        {
            foreach (var kvp in row.Values)
            {
                if (!row.ColumnIndexMap.TryGetValue(kvp.Key, out int zeroBasedCol))
                    continue;

                var cell = worksheet.Cell(row.RowNumber, zeroBasedCol + 1);
                var currentText = cell.GetFormattedString();
                if (currentText != kvp.Value)
                {
                    cell.Value = kvp.Value;
                    updatedCells++;
                }
            }
        }

        if (updatedCells > 0)
        {
            workbook.Save();
            foreach (var row in modifiedRows)
                row.MarkSaved();
        }

        return updatedCells;
    }

    private static int SaveAsDecryptedCopy(
        string sourceFilePath,
        string destinationFilePath,
        string? targetSheetName,
        List<EditableRowMatch> modifiedRows,
        string? password)
    {
        // Map edits by 1-based row number -> (0-based column index -> new value)
        var editsByRow = new Dictionary<int, Dictionary<int, string>>();
        foreach (var row in modifiedRows)
        {
            if (!editsByRow.TryGetValue(row.RowNumber, out var colEdits))
            {
                colEdits = new Dictionary<int, string>();
                editsByRow[row.RowNumber] = colEdits;
            }

            foreach (var kvp in row.Values)
            {
                if (row.ColumnIndexMap.TryGetValue(kvp.Key, out int zeroBasedCol))
                {
                    colEdits[zeroBasedCol] = kvp.Value;
                }
            }
        }

        using var stream = File.Open(sourceFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration
        {
            Password = password
        });

        using var outWorkbook = new XLWorkbook();
        int updatedCells = 0;
        int sheetIndex = 0;

        do
        {
            string currentSheetName = string.IsNullOrWhiteSpace(reader.Name)
                ? $"Sheet{sheetIndex + 1}"
                : reader.Name;

            bool isTargetSheet = string.IsNullOrWhiteSpace(targetSheetName)
                ? sheetIndex == 0
                : string.Equals(currentSheetName, targetSheetName, StringComparison.OrdinalIgnoreCase);

            var outSheet = outWorkbook.Worksheets.Add(currentSheetName);

            while (reader.Read())
            {
                int rowNumber = reader.Depth + 1;
                editsByRow.TryGetValue(rowNumber, out var rowEdits);

                for (int col = 0; col < reader.FieldCount; col++)
                {
                    object? rawValue = reader.GetValue(col);
                    string originalText = rawValue?.ToString() ?? string.Empty;

                    if (isTargetSheet && rowEdits is not null && rowEdits.TryGetValue(col, out var editedValue))
                    {
                        outSheet.Cell(rowNumber, col + 1).Value = editedValue;
                        if (originalText != editedValue)
                            updatedCells++;
                    }
                    else if (rawValue is not null)
                    {
                        outSheet.Cell(rowNumber, col + 1).Value = XLCellValue.FromObject(rawValue);
                    }
                }
            }

            sheetIndex++;
        } while (reader.NextResult());

        if (updatedCells > 0)
        {
            outWorkbook.SaveAs(destinationFilePath);
            foreach (var row in modifiedRows)
                row.MarkSaved();
        }

        return updatedCells;
    }

    private static string BuildEditedCopyPath(string originalFilePath)
    {
        string dir = Path.GetDirectoryName(originalFilePath) ?? Environment.CurrentDirectory;
        string nameWithoutExt = Path.GetFileNameWithoutExtension(originalFilePath);
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        return Path.Combine(dir, $"{nameWithoutExt}_edited_{timestamp}.xlsx");
    }

    private static bool IsEncryptionOrFormatLimitation(Exception ex)
    {
        for (Exception? current = ex; current is not null; current = current.InnerException)
        {
            if (current.Message.Contains("password", StringComparison.OrdinalIgnoreCase) ||
                current.Message.Contains("encrypted", StringComparison.OrdinalIgnoreCase) ||
                current.Message.Contains("corrupt", StringComparison.OrdinalIgnoreCase) ||
                current.Message.Contains("File contains corrupted data", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }
}