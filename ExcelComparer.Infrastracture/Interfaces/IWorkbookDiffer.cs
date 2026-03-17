using DocumentFormat.OpenXml.Packaging;
using ExcelComparer.Application.Models;
using ExcelComparer.Domain.Entities;

namespace ExcelComparer.Infrastructure.Interfaces;

internal interface IWorkbookDiffer
{
    ValueTask DiffSheetsAsync(WorkbookInfo workbookA, WorkbookInfo workbookB, ComparisonResult result);

    ValueTask DiffSheetOrderAsync(SpreadsheetDocument documentA, SpreadsheetDocument documentB, ComparisonResult result);
}
