using ExcelComparer.Application.Interfaces;
using ExcelComparer.Application.Models;
using ExcelComparer.WPFUI.Core;
using ExcelComparer.WPFUI.Models;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Data;

namespace ExcelComparer.WPFUI.ViewModels;

public sealed class MainViewModel : ViewModelBase
{
    private readonly IExcelComparer _excelComparer;
    private readonly RelayCommand _browseFileACommand;
    private readonly RelayCommand _browseFileBCommand;
    private readonly RelayCommand _compareCommand;
    private readonly RelayCommand _cancelCommand;

    private CancellationTokenSource? _cts;
    private string? _selectedSheetFilter;
    private SummaryTreeItem? _selectedSummaryItem;
    private string _fileA = string.Empty;
    private string _fileB = string.Empty;
    private bool _compareValues = true;
    private bool _compareFormulas = true;
    private bool _includeHiddenSheets = true;
    private string _statusText = "Listo.";
    private int _progressValue;
    private string _filterText = string.Empty;
    private bool _isComparing;

    public MainViewModel(IExcelComparer excelComparer)
    {
        _excelComparer = excelComparer;

        DiffsView = CollectionViewSource.GetDefaultView(Diffs);
        DiffsView.Filter = FilterDiffs;
        DiffsView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(DiffItem.Sheet)));

        _browseFileACommand = new RelayCommand(BrowseFileA, () => !IsComparing);
        _browseFileBCommand = new RelayCommand(BrowseFileB, () => !IsComparing);
        _compareCommand = new RelayCommand(async () => await CompareAsync(), () => !IsComparing);
        _cancelCommand = new RelayCommand(CancelComparison, () => IsComparing);
    }

    public ObservableCollection<DiffItem> Diffs { get; } = new ObservableCollection<DiffItem>();

    public ObservableCollection<SummaryTreeItem> SummaryItems { get; } = new ObservableCollection<SummaryTreeItem>();

    public ICollectionView DiffsView { get; }

    public RelayCommand BrowseFileACommand => _browseFileACommand;

    public RelayCommand BrowseFileBCommand => _browseFileBCommand;

    public RelayCommand CompareCommand => _compareCommand;

    public RelayCommand CancelCommand => _cancelCommand;

    public string FileA
    {
        get => _fileA;
        set => SetProperty(ref _fileA, value);
    }

    public string FileB
    {
        get => _fileB;
        set => SetProperty(ref _fileB, value);
    }

    public bool CompareValues
    {
        get => _compareValues;
        set => SetProperty(ref _compareValues, value);
    }

    public bool CompareFormulas
    {
        get => _compareFormulas;
        set => SetProperty(ref _compareFormulas, value);
    }

    public bool IncludeHiddenSheets
    {
        get => _includeHiddenSheets;
        set => SetProperty(ref _includeHiddenSheets, value);
    }

    public string StatusText
    {
        get => _statusText;
        set => SetProperty(ref _statusText, value);
    }

    public int ProgressValue
    {
        get => _progressValue;
        set => SetProperty(ref _progressValue, value);
    }

    public string FilterText
    {
        get => _filterText;
        set
        {
            if (SetProperty(ref _filterText, value))
            {
                DiffsView.Refresh();
            }
        }
    }

    public bool IsComparing
    {
        get => _isComparing;
        private set
        {
            if (SetProperty(ref _isComparing, value))
            {
                _browseFileACommand.NotifyCanExecuteChanged();
                _browseFileBCommand.NotifyCanExecuteChanged();
                _compareCommand.NotifyCanExecuteChanged();
                _cancelCommand.NotifyCanExecuteChanged();
            }
        }
    }

    public SummaryTreeItem? SelectedSummaryItem
    {
        get => _selectedSummaryItem;
        set
        {
            if (SetProperty(ref _selectedSummaryItem, value))
            {
                _selectedSheetFilter = value?.Tag;
                DiffsView.Refresh();
            }
        }
    }

    public override void Dispose()
    {
        _cts?.Dispose();
        _cts = null;
        base.Dispose();
    }

    private void BrowseFileA()
    {
        FileA = PickExcelFile();
    }

    private void BrowseFileB()
    {
        FileB = PickExcelFile();
    }

    private static string PickExcelFile()
    {
        var dlg = new OpenFileDialog
        {
            Filter = "Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*",
            CheckFileExists = true
        };

        return dlg.ShowDialog() == true ? dlg.FileName : string.Empty;
    }

    private async Task CompareAsync()
    {
        var fileA = FileA?.Trim();
        var fileB = FileB?.Trim();

        if (!CanCompareFiles(fileA, fileB))
        {
            MessageBox.Show("Selecciona ambos archivos.", "Excel Diff", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        PrepareForComparison();
        _cts = new CancellationTokenSource();
        var options = CreateComparisonOptions();
        var progress = CreateProgressReporter();

        try
        {
            var result = await RunComparisonAsync(fileA!, fileB!, options, progress, _cts.Token);
            ShowComparisonResult(result);
        }
        catch (OperationCanceledException)
        {
            StatusText = "Cancelado.";
        }
        catch (Exception ex)
        {
            StatusText = "Error.";
            MessageBox.Show(ex.Message, "Excel Diff - Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            IsComparing = false;
            _cts?.Dispose();
            _cts = null;
        }
    }

    private static bool CanCompareFiles(string? fileA, string? fileB)
        => !string.IsNullOrWhiteSpace(fileA) && !string.IsNullOrWhiteSpace(fileB);

    private Task<ComparisonResult> RunComparisonAsync(
        string fileA,
        string fileB,
        ComparisonOptions options,
        IProgress<ProgressInfo> progress,
        CancellationToken cancellationToken)
        => Task.Run(
            async () => await _excelComparer.CompareAsync(fileA, fileB, options, progress, cancellationToken),
            cancellationToken);

    private void PrepareForComparison()
    {
        IsComparing = true;
        StatusText = "Comparando...";
        ProgressValue = 0;
        Diffs.Clear();
        SummaryItems.Clear();
        _selectedSheetFilter = null;
        SelectedSummaryItem = null;
    }

    private ComparisonOptions CreateComparisonOptions()
        => new()
        {
            CompareValues = CompareValues,
            CompareFormulas = CompareFormulas,
            IncludeHiddenSheets = IncludeHiddenSheets
        };

    private IProgress<ProgressInfo> CreateProgressReporter()
        => new Progress<ProgressInfo>(update =>
        {
            ProgressValue = update.Percent;
            StatusText = update.Message;
        });

    private void ShowComparisonResult(ComparisonResult result)
    {
        foreach (var diff in result.Diffs)
        {
            Diffs.Add(diff);
        }

        SummaryItems.Add(BuildSummaryTree(result));
        StatusText = $"Listo. Cambios: {Diffs.Count}";
        ProgressValue = 100;
    }

    private void CancelComparison()
    {
        _cts?.Cancel();
    }

    private bool FilterDiffs(object obj)
    {
        if (obj is not DiffItem diff)
        {
            return false;
        }

        if (!string.IsNullOrEmpty(_selectedSheetFilter)
            && !string.Equals(diff.Sheet, _selectedSheetFilter, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var query = FilterText?.Trim();
        if (string.IsNullOrWhiteSpace(query))
        {
            return true;
        }

        return Contains(diff.Sheet, query)
            || Contains(diff.Address, query)
            || Contains(diff.Kind.ToString(), query)
            || Contains(diff.What, query)
            || Contains(diff.Before, query)
            || Contains(diff.After, query);
    }

    private static bool Contains(string? value, string query)
        => value?.Contains(query, StringComparison.OrdinalIgnoreCase) == true;

    private static SummaryTreeItem BuildSummaryTree(ComparisonResult result)
    {
        var root = new SummaryTreeItem
        {
            Header = $"Workbook diff (total: {result.Diffs.Count})",
            IsExpanded = true
        };

        var bySheet = result.Diffs
            .GroupBy(d => d.Sheet)
            .OrderBy(g => g.Key);

        foreach (var group in bySheet)
        {
            var added = group.Count(x => x.Kind == DiffKind.Added);
            var removed = group.Count(x => x.Kind == DiffKind.Removed);
            var modified = group.Count(x => x.Kind == DiffKind.Modified);

            root.Children.Add(new SummaryTreeItem
            {
                Header = $"{group.Key}   (+{added}  ~{modified}  -{removed})",
                Tag = group.Key
            });
        }

        return root;
    }
}