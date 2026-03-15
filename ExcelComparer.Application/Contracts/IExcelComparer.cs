using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelComparer.Domain;
using ExcelComparer.Domain.Entities;

namespace ExcelComparer.Application.Contracts;

public interface IExcelComparer
{
    ValueTask<ComparisonResult> CompareAsync(string fileA, string fileB, ComparisonOptions options, IProgress<ProgressInfo>? progress, CancellationToken ct);
    WorkbookInfo ReadWorkbook(SpreadsheetDocument doc);
    ValueTask DiffSheets(WorkbookInfo a, WorkbookInfo b, ComparisonResult result, ComparisonOptions options);
    ValueTask DiffSheetOrder(SpreadsheetDocument docA, SpreadsheetDocument docB, ComparisonResult result);
    ValueTask DiffUsedRange(Worksheet wsA, Worksheet wsB, string sheetName, ComparisonResult result);
    string GetUsedRangeFromWorksheet(Worksheet ws);
    void DiffDataValidations(Worksheet wsA, Worksheet wsB, string sheetName, ComparisonResult result);
    string GetDataValidationSummaryFromWorksheet(Worksheet ws);
    ValueTask DiffConditionalFormatting(Worksheet wsA, Worksheet wsB, string sheetName, ComparisonResult result);
    ValueTask DiffHiddenRowsCols(Worksheet wsA, Worksheet wsB, string sheetName, ComparisonResult result);
    ValueTask<Dictionary<string, CellInfo>> ReadCells(WorkbookPart wbPart, Worksheet ws, SharedStringTable? sst, ComparisonOptions options, Dictionary<int, string?>? nfCache);
    ValueTask<string?> ReadCellValueAsText(Cell cell, SharedStringTable? sst);
    ValueTask<Dictionary<int, string?>> BuildNumberFormatCache(WorkbookPart wbPart);
    string? ResolveNumberFormatCode(WorkbookPart wbPart, uint? styleIndex);
    ValueTask DiffCells(string sheetName, Dictionary<string, CellInfo> a, Dictionary<string, CellInfo> b, ComparisonResult result, ComparisonOptions options, WorkbookPart wbA, WorkbookPart wbB);
    string? Summarize(CellInfo? c, ComparisonOptions options);
    ValueTask ParseAddress(string addr, out int row, out int col);
    int ColumnNameToIndex(string name);
    string ColumnIndexToName(int index);
}
