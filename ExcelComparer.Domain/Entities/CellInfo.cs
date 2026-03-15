namespace ExcelComparer.Domain.Entities;

public class CellInfo
{
    public string? ValueText { get; init; }
    public string? FormulaText { get; init; }
    public uint? StyleIndex { get; init; }
    public string? NumberFormatCode { get; init; }
}
