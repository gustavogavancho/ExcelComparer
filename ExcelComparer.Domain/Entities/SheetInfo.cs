namespace ExcelComparer.Domain.Entities;

public class SheetInfo
{
    public required string Name { get; init; }
    public required string SheetId { get; init; }
    public required string RelId { get; init; }
    public bool Hidden { get; init; }
    public bool VeryHidden { get; init; }
}