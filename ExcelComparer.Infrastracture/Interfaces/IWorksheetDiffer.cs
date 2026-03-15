using DocumentFormat.OpenXml.Spreadsheet;
using ExcelComparer.Application.Models;
using ExcelComparer.Domain.Entities;

namespace ExcelComparer.Infrastructure;

internal interface IWorksheetDiffer
{
    ValueTask DiffUsedRangeAsync(Worksheet worksheetA, Worksheet worksheetB, string sheetName, ComparisonResult result);

    void DiffDataValidations(Worksheet worksheetA, Worksheet worksheetB, string sheetName, ComparisonResult result);

    ValueTask DiffConditionalFormattingAsync(Worksheet worksheetA, Worksheet worksheetB, string sheetName, ComparisonResult result);

    ValueTask DiffHiddenRowsColsAsync(Worksheet worksheetA, Worksheet worksheetB, string sheetName, ComparisonResult result);

    ValueTask DiffCellsAsync(
        string sheetName,
        Dictionary<string, CellInfo> cellsA,
        Dictionary<string, CellInfo> cellsB,
        ComparisonResult result,
        ComparisonOptions options);
}
