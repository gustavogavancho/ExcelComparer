namespace ExcelComparer.Domain;

public sealed class ComparisonResult
{
    public List<DiffItem> Diffs { get; }
}