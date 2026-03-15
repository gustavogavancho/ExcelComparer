namespace ExcelComparer.Domain.Entities;

public class WorkbookInfo
{
    public Dictionary<string, SheetInfo> SheetsByName { get; } = new(StringComparer.OrdinalIgnoreCase);
    public List<string> SheetOrder { get; } = new();
}
