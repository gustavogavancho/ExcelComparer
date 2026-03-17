using DocumentFormat.OpenXml.Spreadsheet;
using ExcelComparer.Application.Models;
using ExcelComparer.Domain.Entities;
using ExcelComparer.Infrastructure.Implementations;

namespace ExcelComparer.Infrastructure.UnitTests;

public class WorksheetDifferTests
{
    [Fact]
    public async Task DiffCellsAsync_WhenCellIsRemoved_ShouldAddRemovedDiff()
    {
        var differ = new WorksheetDiffer();
        var result = new ComparisonResult();
        var options = new ComparisonOptions();
        var cellsA = new Dictionary<string, CellInfo>
        {
            ["A1"] = new() { ValueText = "old" }
        };
        var cellsB = new Dictionary<string, CellInfo>(StringComparer.OrdinalIgnoreCase);

        await differ.DiffCellsAsync("Sheet1", cellsA, cellsB, result, options);

        var diff = Assert.Single(result.Diffs);
        Assert.Equal(DiffKind.Removed, diff.Kind);
        Assert.Equal("Cell", diff.What);
        Assert.Equal("A1", diff.Address);
        Assert.Equal("old", diff.Before);
        Assert.Null(diff.After);
    }

    [Fact]
    public async Task DiffUsedRangeAsync_WhenUsedRangeChanges_ShouldAddDiff()
    {
        var differ = new WorksheetDiffer();
        var result = new ComparisonResult();
        var worksheetA = CreateWorksheet("A1");
        var worksheetB = CreateWorksheet("C3");

        await differ.DiffUsedRangeAsync(worksheetA, worksheetB, "Sheet1", result);

        var diff = Assert.Single(result.Diffs);
        Assert.Equal(DiffKind.Modified, diff.Kind);
        Assert.Equal("UsedRange", diff.What);
        Assert.Equal("A1:A1", diff.Before);
        Assert.Equal("C3:C3", diff.After);
    }

    private static Worksheet CreateWorksheet(string cellReference)
    {
        var rowIndex = uint.Parse(new string(cellReference.SkipWhile(char.IsLetter).ToArray()));
        var row = new Row { RowIndex = rowIndex };
        row.Append(new Cell { CellReference = cellReference, CellValue = new CellValue("1") });
        return new Worksheet(new SheetData(row));
    }
}
