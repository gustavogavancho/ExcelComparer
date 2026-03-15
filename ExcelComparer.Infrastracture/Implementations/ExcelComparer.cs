using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelComparer.Application.Contracts;
using ExcelComparer.Application.Models;

namespace ExcelComparer.Infrastructure;

public class ExcelComparer : IExcelComparer
{
    private readonly IOpenXmlWorkbookReader _workbookReader;
    private readonly IWorkbookDiffer _workbookDiffer;
    private readonly IWorksheetDiffer _worksheetDiffer;

    internal ExcelComparer(
        IOpenXmlWorkbookReader workbookReader,
        IWorkbookDiffer workbookDiffer,
        IWorksheetDiffer worksheetDiffer)
    {
        _workbookReader = workbookReader;
        _workbookDiffer = workbookDiffer;
        _worksheetDiffer = worksheetDiffer;
    }

    public async ValueTask<ComparisonResult> CompareAsync(string fileA, string fileB, ComparisonOptions options, IProgress<ProgressInfo>? progress, CancellationToken ct)
    {
        var result = new ComparisonResult();
        progress?.Report(new ProgressInfo(1, "Leyendo estructura de libros..."));

        using var documentA = SpreadsheetDocument.Open(fileA, false);
        using var documentB = SpreadsheetDocument.Open(fileB, false);

        var workbookA = _workbookReader.ReadWorkbook(documentA);
        var workbookB = _workbookReader.ReadWorkbook(documentB);

        progress?.Report(new ProgressInfo(5, "Comparando hojas..."));
        await _workbookDiffer.DiffSheetsAsync(workbookA, workbookB, result);

        if (options.CompareSheetOrder)
        {
            await _workbookDiffer.DiffSheetOrderAsync(documentA, documentB, result);
        }

        var workbookPartA = documentA.WorkbookPart!;
        var workbookPartB = documentB.WorkbookPart!;
        var numberFormatCacheA = await _workbookReader.BuildNumberFormatCacheAsync(workbookPartA);
        var numberFormatCacheB = await _workbookReader.BuildNumberFormatCacheAsync(workbookPartB);

        var allSheetNames = workbookA.SheetsByName.Keys
            .Union(workbookB.SheetsByName.Keys)
            .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();

        for (int i = 0; i < allSheetNames.Count; i++)
        {
            ct.ThrowIfCancellationRequested();

            var sheetName = allSheetNames[i];
            ReportSheetProgress(progress, sheetName, i, allSheetNames.Count);

            if (!TryGetSheetPair(workbookA, workbookB, sheetName, out var sheetA, out var sheetB))
            {
                continue;
            }

            if (ShouldSkipSheet(options, sheetA!, sheetB!))
            {
                continue;
            }

            var worksheetA = ((WorksheetPart)workbookPartA.GetPartById(sheetA!.RelId)).Worksheet;
            var worksheetB = ((WorksheetPart)workbookPartB.GetPartById(sheetB!.RelId)).Worksheet;

            await DiffWorksheetMetadataAsync(options, worksheetA, worksheetB, sheetName, result);

            var cellsA = await _workbookReader.ReadCellsAsync(workbookPartA, worksheetA, workbookPartA.SharedStringTablePart?.SharedStringTable, options, numberFormatCacheA);
            var cellsB = await _workbookReader.ReadCellsAsync(workbookPartB, worksheetB, workbookPartB.SharedStringTablePart?.SharedStringTable, options, numberFormatCacheB);

            await _worksheetDiffer.DiffCellsAsync(sheetName, cellsA, cellsB, result, options);
        }

        progress?.Report(new ProgressInfo(100, "Finalizado."));
        return result;
    }

    private async ValueTask DiffWorksheetMetadataAsync(ComparisonOptions options, Worksheet worksheetA, Worksheet worksheetB, string sheetName, ComparisonResult result)
    {
        if (options.CompareUsedRange)
        {
            await _worksheetDiffer.DiffUsedRangeAsync(worksheetA, worksheetB, sheetName, result);
        }

        if (options.CompareValidations)
        {
            _worksheetDiffer.DiffDataValidations(worksheetA, worksheetB, sheetName, result);
        }

        if (options.CompareConditionalFormats)
        {
            await _worksheetDiffer.DiffConditionalFormattingAsync(worksheetA, worksheetB, sheetName, result);
        }

        if (options.CompareHiddenRowsCols)
        {
            await _worksheetDiffer.DiffHiddenRowsColsAsync(worksheetA, worksheetB, sheetName, result);
        }
    }

    private static bool TryGetSheetPair(
        Domain.Entities.WorkbookInfo workbookA,
        Domain.Entities.WorkbookInfo workbookB,
        string sheetName,
        out Domain.Entities.SheetInfo? sheetA,
        out Domain.Entities.SheetInfo? sheetB)
    {
        var hasA = workbookA.SheetsByName.TryGetValue(sheetName, out sheetA);
        var hasB = workbookB.SheetsByName.TryGetValue(sheetName, out sheetB);
        return hasA && hasB;
    }

    private static bool ShouldSkipSheet(ComparisonOptions options, Domain.Entities.SheetInfo sheetA, Domain.Entities.SheetInfo sheetB)
        => !options.IncludeHiddenSheets && (sheetA.Hidden || sheetB.Hidden);

    private static void ReportSheetProgress(IProgress<ProgressInfo>? progress, string sheetName, int currentIndex, int totalSheets)
    {
        var percentage = 10 + (int)((currentIndex / (double)Math.Max(1, totalSheets)) * 85);
        progress?.Report(new ProgressInfo(percentage, $"Comparando hoja: {sheetName} ({currentIndex + 1}/{totalSheets})"));
    }
}