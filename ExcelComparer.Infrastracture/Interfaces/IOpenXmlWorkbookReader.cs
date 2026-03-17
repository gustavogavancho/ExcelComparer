using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelComparer.Application.Models;
using ExcelComparer.Domain.Entities;

namespace ExcelComparer.Infrastructure.Interfaces;

internal interface IOpenXmlWorkbookReader
{
    ValueTask<Dictionary<int, string?>> BuildNumberFormatCacheAsync(WorkbookPart workbookPart);

    WorkbookInfo ReadWorkbook(SpreadsheetDocument document);

    ValueTask<Dictionary<string, CellInfo>> ReadCellsAsync(
        WorkbookPart workbookPart,
        Worksheet worksheet,
        SharedStringTable? sharedStringTable,
        ComparisonOptions options,
        Dictionary<int, string?>? numberFormatCache);
}
