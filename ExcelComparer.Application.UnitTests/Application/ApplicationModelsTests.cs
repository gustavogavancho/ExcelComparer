using ExcelComparer.Application.Models;

namespace ExcelComparer.Application.UnitTests.Application;

public class ApplicationModelsTests
{
    [Fact]
    public void ComparisonOptions_ShouldExposeExpectedDefaults()
    {
        var options = new ComparisonOptions();

        Assert.True(options.CompareValues);
        Assert.True(options.CompareFormulas);
        Assert.True(options.IncludeHiddenSheets);
        Assert.True(options.CompareSheetOrder);
        Assert.False(options.CompareWorkbookProperties);
        Assert.True(options.CompareUsedRange);
        Assert.True(options.CompareValidations);
        Assert.False(options.CompareConditionalFormats);
        Assert.False(options.CompareHiddenRowsCols);
        Assert.False(options.CompareCellFormat);
    }

    [Fact]
    public void ComparisonResult_ShouldStartWithEmptyDiffCollection()
    {
        var result = new ComparisonResult();

        Assert.NotNull(result.Diffs);
        Assert.Empty(result.Diffs);
    }

    [Fact]
    public void DiffItem_And_ProgressInfo_ShouldPreserveAssignedValues()
    {
        var diff = new DiffItem("Sheet1", "A1", DiffKind.Modified, "Value", "old", "new");
        var progress = new ProgressInfo(50, "Comparando");

        Assert.Equal("Sheet1", diff.Sheet);
        Assert.Equal("A1", diff.Address);
        Assert.Equal(DiffKind.Modified, diff.Kind);
        Assert.Equal("Value", diff.What);
        Assert.Equal("old", diff.Before);
        Assert.Equal("new", diff.After);
        Assert.Equal(50, progress.Percent);
        Assert.Equal("Comparando", progress.Message);
    }
}
