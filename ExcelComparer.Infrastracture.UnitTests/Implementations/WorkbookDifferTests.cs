using ExcelComparer.Application.Models;
using ExcelComparer.Domain.Entities;

namespace ExcelComparer.Infrastructure.UnitTests;

public class WorkbookDifferTests
{
    [Fact]
    public async Task DiffSheetsAsync_WhenVisibilityChanges_ShouldAddSheetVisibilityDiff()
    {
        var differ = new WorkbookDiffer();
        var result = new ComparisonResult();
        var workbookA = CreateWorkbookInfo(hidden: false);
        var workbookB = CreateWorkbookInfo(hidden: true);

        await differ.DiffSheetsAsync(workbookA, workbookB, result);

        var diff = Assert.Single(result.Diffs);
        Assert.Equal("Sheet1", diff.Sheet);
        Assert.Equal(DiffKind.Modified, diff.Kind);
        Assert.Equal("SheetVisibility", diff.What);
        Assert.Equal("Visible", diff.Before);
        Assert.Equal("Hidden", diff.After);
    }

    private static WorkbookInfo CreateWorkbookInfo(bool hidden)
    {
        var workbook = new WorkbookInfo();
        workbook.SheetsByName["Sheet1"] = new SheetInfo
        {
            Name = "Sheet1",
            SheetId = "1",
            RelId = "rId1",
            Hidden = hidden,
            VeryHidden = false
        };

        return workbook;
    }
}
