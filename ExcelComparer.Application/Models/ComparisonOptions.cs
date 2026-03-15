namespace ExcelComparer.Application.Models;

public sealed class ComparisonOptions
{
    public bool CompareValues { get; set; } = true;
    public bool CompareFormulas { get; set; } = true;
    public bool IncludeHiddenSheets { get; set; } = true;

    public bool CompareSheetOrder { get; set; } = true;
    public bool CompareWorkbookProperties { get; set; } = false;

    public bool CompareUsedRange { get; set; } = true;
    public bool CompareValidations { get; set; } = true;
    public bool CompareConditionalFormats { get; set; } = false;
    public bool CompareHiddenRowsCols { get; set; } = false;

    public bool CompareCellFormat { get; set; } = false;
}
