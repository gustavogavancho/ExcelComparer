namespace ExcelComparer.Application.Models;

public sealed class ComparisonResult
{
    public List<DiffItem> Diffs { get; } = [];
}
