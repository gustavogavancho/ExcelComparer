using System.Collections.ObjectModel;

namespace ExcelComparer.WPFUI.Models;

public sealed class SummaryTreeItem
{
    public string Header { get; init; } = string.Empty;

    public string? Tag { get; init; }

    public bool IsExpanded { get; init; }

    public ObservableCollection<SummaryTreeItem> Children { get; } = [];
}
