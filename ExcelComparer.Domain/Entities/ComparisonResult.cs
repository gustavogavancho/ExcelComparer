using ExcelComparer.Domain.Enums;

namespace ExcelComparer.Domain.Entities;

public sealed class ComparisonResult
{
    public List<DiffItem> Diffs { get; }
}