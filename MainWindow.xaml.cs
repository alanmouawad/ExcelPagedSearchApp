using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;
using ExcelDataReader.Exceptions;
using ExcelPagedSearchApp.Models;
using ExcelPagedSearchApp.Services;
using ExcelPagedSearchApp.Views;
using Microsoft.Win32;

namespace ExcelPagedSearchApp;

public partial class MainWindow : Window
{
    private readonly ExcelPagedSearcher _searcher = new();
    private readonly ExcelWriterService _writer = new();
    private readonly ObservableCollection<EditableRowMatch> _matches = new();
    private readonly ObservableCollection<ColumnOption> _searchColumns = new();
    private readonly ObservableCollection<ColumnOption> _displayColumns = new();

    private readonly DispatcherTimer _uiTimer;
    private readonly Stopwatch _stopwatch = new();
    private string _activeOperationLabel = "Ready";

    private ICollectionView? _searchColumnsView;
    private ICollectionView? _displayColumnsView;
    private CancellationTokenSource? _cts;
    private bool _isLoadingColumns;
    private int _totalMatches;
    private int _lastPageSeen;
    private string? _currentFilePassword;

    public MainWindow()
    {
        InitializeComponent();
        ResultsGrid.ItemsSource = _matches;

        _searchColumnsView = CollectionViewSource.GetDefaultView(_searchColumns);
        _searchColumnsView.Filter = FilterSearchColumn;
        SearchColumnsListBox.ItemsSource = _searchColumnsView;

        _displayColumnsView = CollectionViewSource.GetDefaultView(_displayColumns);
        _displayColumnsView.Filter = FilterDisplayColumn;
        DisplayColumnsListBox.ItemsSource = _displayColumnsView;

        _uiTimer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(250)
        };
        _uiTimer.Tick += UiTimer_Tick;

        UpdateColumnCountHeaders();
    }

    private void StartElapsedCounter(string operationLabel)
    {
        _activeOperationLabel = operationLabel;
        _stopwatch.Restart();
        _uiTimer.Start();
        UpdateElapsedText();
    }

    private void StopElapsedCounter()
    {
        _uiTimer.Stop();
        _stopwatch.Stop();
        UpdateElapsedText();
    }

    private void UiTimer_Tick(object? sender, EventArgs e)
    {
        UpdateElapsedText();
        if (_cts is not null && !_cts.IsCancellationRequested && _uiTimer.IsEnabled)
        {
            StatusTextBlock.Text = $"{_activeOperationLabel} — Page {_lastPageSeen}, {_totalMatches} match(es) ({FormatSeconds(_stopwatch.Elapsed)})";
        }
    }

    private void UpdateElapsedText()
        => ElapsedTimerTextBlock.Text = $"Elapsed: {FormatSeconds(_stopwatch.Elapsed)}";

    private static string FormatSeconds(TimeSpan ts)
        => ts.TotalSeconds < 10 ? $"{ts.TotalSeconds:0.0}s" : $"{(int)ts.TotalSeconds}s";

    private bool FilterSearchColumn(object item)
    {
        if (item is not ColumnOption col) return false;
        var filter = SearchColFilterTextBox.Text?.Trim();
        return string.IsNullOrEmpty(filter) ||
               col.HeaderName.Contains(filter, StringComparison.OrdinalIgnoreCase);
    }

    private bool FilterDisplayColumn(object item)
    {
        if (item is not ColumnOption col) return false;
        var filter = DisplayColFilterTextBox.Text?.Trim();
        return string.IsNullOrEmpty(filter) ||
               col.HeaderName.Contains(filter, StringComparison.OrdinalIgnoreCase);
    }

    private void SearchColFilterTextBox_TextChanged(object sender, TextChangedEventArgs e)
        => _searchColumnsView?.Refresh();

    private void DisplayColFilterTextBox_TextChanged(object sender, TextChangedEventArgs e)
        => _displayColumnsView?.Refresh();

    private void ColumnSelection_Changed(object sender, RoutedEventArgs e)
        => UpdateColumnCountHeaders();

    private void UpdateColumnCountHeaders()
    {
        int selectedSearch = _searchColumns.Count(c => c.IsSelected);
        int selectedDisplay = _displayColumns.Count(c => c.IsSelected);

        SearchColumnsHeaderText.Text = _searchColumns.Count > 0
            ? $"Search In Columns ({selectedSearch}/{_searchColumns.Count})"
            : "Search In Columns";

        DisplayColumnsHeaderText.Text = _displayColumns.Count > 0
            ? $"Show In Results Columns ({selectedDisplay}/{_displayColumns.Count})"
            : "Show In Results Columns";
    }

    private void BrowseButton_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*"
        };

        if (dialog.ShowDialog() != true)
            return;

        _currentFilePassword = null;
        StartElapsedCounter("Loading sheets & columns");

        try
        {
            FilePathTextBox.Text = dialog.FileName;
            IReadOnlyList<string> sheets;

            while (true)
            {
                try
                {
                    sheets = ExcelPagedSearcher.GetSheetNames(dialog.FileName, _currentFilePassword);
                    break;
                }
                catch (Exception ex) when (IsEncryptionOrPasswordException(ex))
                {
                    string promptMsg = _currentFilePassword is null
                        ? $"'{Path.GetFileName(dialog.FileName)}' is encrypted. Enter the password:"
                        : "Incorrect password. Please try again:";

                    var pwd = PromptForPassword(dialog.FileName, promptMsg);
                    if (pwd is null)
                    {
                        StopElapsedCounter();
                        StatusTextBlock.Text = "Encrypted file open cancelled.";
                        return;
                    }
                    _currentFilePassword = pwd;
                }
            }

            SheetComboBox.ItemsSource = sheets;

            if (sheets.Count > 0)
            {
                SheetComboBox.SelectedIndex = 0;
                LoadColumnsForSelectedSheet();
            }

            StopElapsedCounter();
            StatusTextBlock.Text = $"Loaded '{Path.GetFileName(dialog.FileName)}' ({sheets.Count} sheet(s)) in {FormatSeconds(_stopwatch.Elapsed)}.";
        }
        catch (Exception ex)
        {
            StopElapsedCounter();
            MessageBox.Show(ex.Message, "Load Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void SheetComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (!_isLoadingColumns)
            LoadColumnsForSelectedSheet();
    }

    private void HeaderOption_Changed(object sender, RoutedEventArgs e)
    {
        if (!_isLoadingColumns)
            LoadColumnsForSelectedSheet();
    }

    private void LoadColumnsForSelectedSheet()
    {
        var filePath = FilePathTextBox.Text;
        var sheetName = SheetComboBox.SelectedItem as string;
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath) || string.IsNullOrWhiteSpace(sheetName))
            return;

        _isLoadingColumns = true;
        StartElapsedCounter("Reading sheet columns");
        try
        {
            IReadOnlyList<ColumnOption> headers;
            while (true)
            {
                try
                {
                    headers = ExcelPagedSearcher.GetColumnHeaders(
                        filePath,
                        sheetName,
                        HasHeaderCheckBox.IsChecked == true,
                        _currentFilePassword);
                    break;
                }
                catch (Exception ex) when (IsEncryptionOrPasswordException(ex))
                {
                    var pwd = PromptForPassword(filePath, "Password required or incorrect. Enter the workbook password:");
                    if (pwd is null)
                    {
                        StopElapsedCounter();
                        StatusTextBlock.Text = "Column load cancelled (file is password-protected).";
                        return;
                    }
                    _currentFilePassword = pwd;
                }
            }

            _searchColumns.Clear();
            _displayColumns.Clear();
            SearchColFilterTextBox.Text = string.Empty;
            DisplayColFilterTextBox.Text = string.Empty;

            foreach (var h in headers)
            {
                _searchColumns.Add(new ColumnOption { Index = h.Index, HeaderName = h.HeaderName, IsSelected = true });
                _displayColumns.Add(new ColumnOption { Index = h.Index, HeaderName = h.HeaderName, IsSelected = true });
            }

            _searchColumnsView?.Refresh();
            _displayColumnsView?.Refresh();
            UpdateColumnCountHeaders();
            StopElapsedCounter();
            StatusTextBlock.Text = $"Loaded {headers.Count} column(s) from '{sheetName}' in {FormatSeconds(_stopwatch.Elapsed)}.";
        }
        finally
        {
            _isLoadingColumns = false;
        }
    }

    private void SelectAllSearchCols_Click(object sender, RoutedEventArgs e)
    {
        foreach (var col in _searchColumnsView?.Cast<ColumnOption>() ?? _searchColumns)
            col.IsSelected = true;
        UpdateColumnCountHeaders();
    }

    private void ClearSearchCols_Click(object sender, RoutedEventArgs e)
    {
        foreach (var col in _searchColumnsView?.Cast<ColumnOption>() ?? _searchColumns)
            col.IsSelected = false;
        UpdateColumnCountHeaders();
    }

    private void SelectAllDisplayCols_Click(object sender, RoutedEventArgs e)
    {
        foreach (var col in _displayColumnsView?.Cast<ColumnOption>() ?? _displayColumns)
            col.IsSelected = true;
        UpdateColumnCountHeaders();
    }

    private void ClearDisplayCols_Click(object sender, RoutedEventArgs e)
    {
        foreach (var col in _displayColumnsView?.Cast<ColumnOption>() ?? _displayColumns)
            col.IsSelected = false;
        UpdateColumnCountHeaders();
    }

    private async void SearchButton_Click(object sender, RoutedEventArgs e)
        => await RunPagedSearchAsync();

    private async void QueryTextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
            await RunPagedSearchAsync();
    }

    private void CancelButton_Click(object sender, RoutedEventArgs e)
        => _cts?.Cancel();

    private async Task RunPagedSearchAsync()
    {
        var filePath = FilePathTextBox.Text;
        var query = QueryTextBox.Text;

        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            MessageBox.Show("Please choose a valid Excel file first.", "Missing File", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(query))
        {
            MessageBox.Show("Please enter a search term.", "Missing Query", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!int.TryParse(PageSizeTextBox.Text, out int pageSize) || pageSize <= 0)
        {
            MessageBox.Show("Rows per page must be a positive number.", "Invalid Page Size", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!int.TryParse(PageStrideTextBox.Text, out int pageStride) || pageStride <= 0)
        {
            pageStride = 1;
            PageStrideTextBox.Text = "1";
        }

        var searchCols = _searchColumns.Where(c => c.IsSelected).Select(c => c.Index).ToHashSet();
        var selectedDisplayColumnOptions = _displayColumns.Where(c => c.IsSelected).ToList();
        var displayCols = selectedDisplayColumnOptions.Select(c => c.Index).ToHashSet();

        if (searchCols.Count == 0)
        {
            MessageBox.Show("Select at least one column to search in.", "No Search Columns", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (displayCols.Count == 0)
        {
            MessageBox.Show("Select at least one column to display in results.", "No Display Columns", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        RebuildResultGridColumns(selectedDisplayColumnOptions);

        _cts?.Cancel();
        _cts = new CancellationTokenSource();

        SetSearchingState(true);
        StartElapsedCounter($"Searching '{query}' (page size {pageSize}, stride {pageStride})");
        StatusTextBlock.Text = $"Searching '{query}' across {searchCols.Count} column(s) in pages of {pageSize} rows...";

        var options = new PageSearchOptions
        {
            PageSize = pageSize,
            HasHeaderRow = HasHeaderCheckBox.IsChecked == true,
            CaseSensitive = CaseSensitiveCheckBox.IsChecked == true,
            StopOnFirstMatch = StopOnFirstMatchCheckBox.IsChecked == true,
            SearchColumnIndices = searchCols,
            DisplayColumnIndices = displayCols
        };

        var selectedSheet = SheetComboBox.SelectedItem as string;

        try
        {
            bool retrySearch;
            do
            {
                retrySearch = false;
                _matches.Clear();
                _totalMatches = 0;
                _lastPageSeen = 0;

                try
                {
                    var stream = pageStride > 1
                        ? _searcher.SearchDistributedSynchronizedAsync(filePath, selectedSheet, query, options, pageStride, _currentFilePassword, _cts.Token)
                        : _searcher.SearchByPagesAsync(filePath, selectedSheet, query, options, _currentFilePassword, _cts.Token);

                    await foreach (var pageResult in stream)
                    {
                        _lastPageSeen = pageResult.PageNumber;
                        foreach (var match in pageResult.Matches)
                        {
                            _matches.Add(match);
                            _totalMatches++;
                        }

                        StatusTextBlock.Text = $"Scanned Page {_lastPageSeen} — {_totalMatches} match(es) found ({FormatSeconds(_stopwatch.Elapsed)}).";
                    }

                    StopElapsedCounter();
                    StatusTextBlock.Text = $"Completed (last page {_lastPageSeen}). Total matches: {_totalMatches} in {FormatSeconds(_stopwatch.Elapsed)}.";
                }
                catch (OperationCanceledException)
                {
                    StopElapsedCounter();
                    StatusTextBlock.Text = $"Search cancelled at Page {_lastPageSeen}. Matches found: {_totalMatches} ({FormatSeconds(_stopwatch.Elapsed)}).";
                }
                catch (Exception ex) when (IsEncryptionOrPasswordException(ex))
                {
                    var pwd = PromptForPassword(filePath, "Password required or incorrect. Enter the workbook password:");
                    if (pwd is not null)
                    {
                        _currentFilePassword = pwd;
                        retrySearch = true;
                    }
                    else
                    {
                        StopElapsedCounter();
                        StatusTextBlock.Text = "Search cancelled (file is password-protected).";
                    }
                }
                catch (Exception ex)
                {
                    StopElapsedCounter();
                    MessageBox.Show(ex.Message, "Search Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    StatusTextBlock.Text = "Search failed.";
                }
            } while (retrySearch);
        }
        finally
        {
            SetSearchingState(false);
            _cts?.Dispose();
            _cts = null;
        }
    }

    private void RebuildResultGridColumns(IEnumerable<ColumnOption> selectedDisplayColumns)
    {
        ResultsGrid.Columns.Clear();

        ResultsGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Page",
            Binding = new Binding(nameof(EditableRowMatch.PageNumber)),
            IsReadOnly = true,
            Width = new DataGridLength(65)
        });

        ResultsGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Row",
            Binding = new Binding(nameof(EditableRowMatch.RowNumber)),
            IsReadOnly = true,
            Width = new DataGridLength(65)
        });

        ResultsGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Matched In",
            Binding = new Binding(nameof(EditableRowMatch.MatchedColumns)),
            IsReadOnly = true,
            Width = new DataGridLength(150)
        });

        ResultsGrid.Columns.Add(new DataGridTextColumn
        {
            Header = "Modified",
            Binding = new Binding(nameof(EditableRowMatch.IsModified)),
            IsReadOnly = true,
            Width = new DataGridLength(75)
        });

        foreach (var col in selectedDisplayColumns)
        {
            ResultsGrid.Columns.Add(new DataGridTextColumn
            {
                Header = col.HeaderName,
                Binding = new Binding($"[{col.HeaderName}]")
                {
                    Mode = BindingMode.TwoWay,
                    UpdateSourceTrigger = UpdateSourceTrigger.LostFocus
                },
                IsReadOnly = false,
                Width = new DataGridLength(160)
            });
        }
    }

    private void SaveChangesButton_Click(object sender, RoutedEventArgs e)
    {
        var filePath = FilePathTextBox.Text;
        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            MessageBox.Show("Please choose a valid Excel file first.", "Missing File", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        StartElapsedCounter("Saving changes to Excel");
        try
        {
            ResultsGrid.CommitEdit(DataGridEditingUnit.Row, exitEditingMode: true);

            var selectedSheet = SheetComboBox.SelectedItem as string;
            var result = _writer.SaveModifiedRows(filePath, selectedSheet, _matches, _currentFilePassword);
            StopElapsedCounter();

            if (result.UpdatedCells == 0)
            {
                StatusTextBlock.Text = $"No modified cells to save ({FormatSeconds(_stopwatch.Elapsed)}).";
                return;
            }

            ResultsGrid.Items.Refresh();

            if (result.SavedAsNewCopy)
            {
                StatusTextBlock.Text = $"Saved {result.UpdatedCells} cell(s) to unlocked copy '{Path.GetFileName(result.SavedFilePath)}' in {FormatSeconds(_stopwatch.Elapsed)}.";
                MessageBox.Show(
                    $"Because the original file is password-protected, {result.UpdatedCells} updated cell(s) were saved to a new unencrypted copy:\n\n{result.SavedFilePath}",
                    "Saved to New Copy",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            else
            {
                StatusTextBlock.Text = $"Saved {result.UpdatedCells} updated cell(s) to '{Path.GetFileName(filePath)}' in {FormatSeconds(_stopwatch.Elapsed)}.";
                MessageBox.Show($"Saved {result.UpdatedCells} updated cell(s) to the Excel file.", "Saved", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            StopElapsedCounter();
            MessageBox.Show(
                $"Could not save changes to Excel. Close the workbook in Excel first and try again.\n\nDetails: {ex.Message}",
                "Save Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }

    private void CopyCell_Click(object sender, RoutedEventArgs e)
    {
        var currentCell = ResultsGrid.CurrentCell;
        if (currentCell.Item is EditableRowMatch row && currentCell.Column is DataGridBoundColumn col)
        {
            var header = col.Header?.ToString();
            string text = !string.IsNullOrEmpty(header) && row.Values.TryGetValue(header, out var val)
                ? val
                : GetSpecialColumnValue(row, header);

            Clipboard.SetText(text ?? string.Empty);
            StatusTextBlock.Text = "Copied cell value to clipboard.";
        }
    }

    private void CopyRow_Click(object sender, RoutedEventArgs e)
    {
        EditableRowMatch? row = ResultsGrid.CurrentItem as EditableRowMatch
            ?? ResultsGrid.SelectedItems.OfType<EditableRowMatch>().FirstOrDefault();

        if (row is null) return;

        string line = FormatRowAsTabLine(row);
        Clipboard.SetText(line);
        StatusTextBlock.Text = $"Copied Row {row.RowNumber} to clipboard.";
    }

    private void CopySelected_Click(object sender, RoutedEventArgs e)
    {
        if (ResultsGrid.SelectedCells.Count > 0)
        {
            var rowsGrouped = ResultsGrid.SelectedCells
                .Where(c => c.Item is EditableRowMatch)
                .GroupBy(c => (EditableRowMatch)c.Item);

            var sb = new StringBuilder();
            foreach (var grp in rowsGrouped)
            {
                var cellTexts = grp
                    .Select(c =>
                    {
                        var h = c.Column?.Header?.ToString();
                        return !string.IsNullOrEmpty(h) && grp.Key.Values.TryGetValue(h, out var v)
                            ? v
                            : GetSpecialColumnValue(grp.Key, h);
                    });
                sb.AppendLine(string.Join("\t", cellTexts));
            }

            if (sb.Length > 0)
            {
                Clipboard.SetText(sb.ToString().TrimEnd());
                StatusTextBlock.Text = "Copied selected cells to clipboard.";
                return;
            }
        }

        var selectedRows = ResultsGrid.SelectedItems.OfType<EditableRowMatch>().ToList();
        if (selectedRows.Count > 0)
        {
            Clipboard.SetText(string.Join(Environment.NewLine, selectedRows.Select(FormatRowAsTabLine)));
            StatusTextBlock.Text = $"Copied {selectedRows.Count} row(s) to clipboard.";
        }
    }

    private static string GetSpecialColumnValue(EditableRowMatch row, string? header) => header switch
    {
        "Page" => row.PageNumber.ToString(),
        "Row" => row.RowNumber.ToString(),
        "Matched In" => row.MatchedColumns,
        "Modified" => row.IsModified.ToString(),
        _ => string.Empty
    };

    private static string FormatRowAsTabLine(EditableRowMatch row)
    {
        var parts = new List<string>
        {
            row.PageNumber.ToString(),
            row.RowNumber.ToString(),
            row.MatchedColumns
        };
        parts.AddRange(row.Values.Values);
        return string.Join("\t", parts);
    }

    private void SetSearchingState(bool isSearching)
    {
        SearchButton.IsEnabled = !isSearching;
        SaveChangesButton.IsEnabled = !isSearching;
        CancelButton.IsEnabled = isSearching;
        BrowseButton.IsEnabled = !isSearching;
        SheetComboBox.IsEnabled = !isSearching;
        SearchColumnsListBox.IsEnabled = !isSearching;
        DisplayColumnsListBox.IsEnabled = !isSearching;
        SearchColFilterTextBox.IsEnabled = !isSearching;
        DisplayColFilterTextBox.IsEnabled = !isSearching;
        QueryTextBox.IsEnabled = !isSearching;
        PageSizeTextBox.IsEnabled = !isSearching;
        PageStrideTextBox.IsEnabled = !isSearching;
    }

    private static bool IsEncryptionOrPasswordException(Exception ex)
{
    for (Exception? current = ex; current is not null; current = current.InnerException)
    {
        if (current.GetType().Name.Contains("Encrypt", StringComparison.OrdinalIgnoreCase) ||
            current.GetType().Name.Contains("Password", StringComparison.OrdinalIgnoreCase) ||
            current.Message.Contains("password", StringComparison.OrdinalIgnoreCase) ||
            current.Message.Contains("encrypted", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
    }

    return false;
}

    private string? PromptForPassword(string filePath, string? customMessage = null)
    {
        bool wasRunning = _uiTimer.IsEnabled;
        if (wasRunning)
            _uiTimer.Stop();

        try
        {
            var dialog = new PasswordPromptWindow(Path.GetFileName(filePath), customMessage)
            {
                Owner = this
            };

            bool? result = dialog.ShowDialog();
            return result == true ? dialog.EnteredPassword : null;
        }
        finally
        {
            if (wasRunning)
                _uiTimer.Start();
        }
    }
}