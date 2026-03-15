using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelComparer.Application.Models;
using ExcelComparer.Domain.Entities;
using System.Globalization;

namespace ExcelComparer.Infrastructure;

internal sealed class WorkbookDiffer : IWorkbookDiffer
{
    private const string EmptyAddress = "";

    public ValueTask DiffSheetsAsync(WorkbookInfo workbookA, WorkbookInfo workbookB, ComparisonResult result)
    {
        var sheetNames = workbookA.SheetsByName.Keys.Union(workbookB.SheetsByName.Keys, StringComparer.OrdinalIgnoreCase);

        foreach (var sheetName in sheetNames.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            var hasA = workbookA.SheetsByName.TryGetValue(sheetName, out var sheetA);
            var hasB = workbookB.SheetsByName.TryGetValue(sheetName, out var sheetB);

            if (hasA && !hasB)
            {
                result.Diffs.Add(new DiffItem(sheetName, EmptyAddress, DiffKind.Removed, "Sheet", "Present", "Missing"));
                continue;
            }

            if (!hasA && hasB)
            {
                result.Diffs.Add(new DiffItem(sheetName, EmptyAddress, DiffKind.Added, "Sheet", "Missing", "Present"));
                continue;
            }

            if (sheetA!.Hidden != sheetB!.Hidden || sheetA.VeryHidden != sheetB.VeryHidden)
            {
                var before = GetVisibilityName(sheetA.Hidden, sheetA.VeryHidden);
                var after = GetVisibilityName(sheetB!.Hidden, sheetB.VeryHidden);
                result.Diffs.Add(new DiffItem(sheetName, EmptyAddress, DiffKind.Modified, "SheetVisibility", before, after));
            }
        }

        return ValueTask.CompletedTask;
    }

    public ValueTask DiffSheetOrderAsync(SpreadsheetDocument documentA, SpreadsheetDocument documentB, ComparisonResult result)
    {
        var orderA = ReadSheetOrder(documentA);
        var orderB = ReadSheetOrder(documentB);
        var indexA = orderA.Select((name, index) => (name, index)).ToDictionary(x => x.name, x => x.index, StringComparer.OrdinalIgnoreCase);
        var indexB = orderB.Select((name, index) => (name, index)).ToDictionary(x => x.name, x => x.index, StringComparer.OrdinalIgnoreCase);

        foreach (var sheetName in indexA.Keys.Intersect(indexB.Keys, StringComparer.OrdinalIgnoreCase))
        {
            var positionA = indexA[sheetName];
            var positionB = indexB[sheetName];
            if (positionA == positionB)
            {
                continue;
            }

            result.Diffs.Add(new DiffItem(
                sheetName,
                EmptyAddress,
                DiffKind.Modified,
                "SheetOrderIndex",
                positionA.ToString(CultureInfo.InvariantCulture),
                positionB.ToString(CultureInfo.InvariantCulture)));
        }

        return ValueTask.CompletedTask;
    }

    private static List<string> ReadSheetOrder(SpreadsheetDocument document)
        => document.WorkbookPart!.Workbook!.Sheets!.OfType<Sheet>().Select(s => s.Name!.Value!).ToList();

    private static string GetVisibilityName(bool isHidden, bool isVeryHidden)
        => isVeryHidden ? "VeryHidden" : (isHidden ? "Hidden" : "Visible");
}
