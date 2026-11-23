using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace ExcelPagedSearchApp.Models;

public sealed class ColumnOption : INotifyPropertyChanged
{
    private bool _isSelected = true;

    public int Index { get; init; }
    public string HeaderName { get; init; } = string.Empty;

    public bool IsSelected
    {
        get => _isSelected;
        set
        {
            if (_isSelected == value) return;
            _isSelected = value;
            OnPropertyChanged();
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}

public sealed class EditableRowMatch : INotifyPropertyChanged
{
    private readonly Dictionary<string, string> _values;
    private readonly Dictionary<string, int> _columnIndexMap;
    private bool _isModified;

    public EditableRowMatch(
        int pageNumber,
        int rowNumber,
        string matchedColumns,
        Dictionary<string, string> values,
        Dictionary<string, int> columnIndexMap)
    {
        PageNumber = pageNumber;
        RowNumber = rowNumber;
        MatchedColumns = matchedColumns;
        _values = values;
        _columnIndexMap = columnIndexMap;
    }

    public int PageNumber { get; }
    public int RowNumber { get; }
    public string MatchedColumns { get; }

    public IReadOnlyDictionary<string, string> Values => _values;
    public IReadOnlyDictionary<string, int> ColumnIndexMap => _columnIndexMap;

    public bool IsModified
    {
        get => _isModified;
        private set
        {
            if (_isModified == value) return;
            _isModified = value;
            OnPropertyChanged();
        }
    }

    public string this[string columnHeader]
    {
        get => _values.TryGetValue(columnHeader, out var val) ? val : string.Empty;
        set
        {
            var newVal = value ?? string.Empty;
            if (_values.TryGetValue(columnHeader, out var current) && current == newVal)
                return;

            _values[columnHeader] = newVal;
            IsModified = true;
            OnPropertyChanged("Item[]");
            OnPropertyChanged(nameof(IsModified));
        }
    }

    public void MarkSaved() => IsModified = false;

    public event PropertyChangedEventHandler? PropertyChanged;

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
}

public sealed class PageSearchOptions
{
    public int PageSize { get; init; } = 500;
    public bool HasHeaderRow { get; init; } = true;
    public bool CaseSensitive { get; init; } = false;
    public bool StopOnFirstMatch { get; init; } = false;
    public IReadOnlySet<int>? SearchColumnIndices { get; init; }
    public IReadOnlySet<int>? DisplayColumnIndices { get; init; }
}

public sealed record PageSearchResult(int PageNumber, IReadOnlyList<EditableRowMatch> Matches);