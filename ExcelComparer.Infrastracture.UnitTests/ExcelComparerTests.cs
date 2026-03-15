using ExcelComparer.Application.Models;
using static ExcelComparer.Infrastracture.UnitTests.TestWorkbookFactory;

namespace ExcelComparer.Infrastracture.UnitTests;

public class ExcelComparerTests
{
    [Fact]
    public async Task CompareAsync_WhenWorkbookBContainsNewSheet_ShouldReturnAddedSheetDiff()
    {
        var comparer = new ExcelComparer.Infrastracture.ExcelComparer();
        var fileA = CreateWorkbook(new TestSheet("Sheet1", values: new Dictionary<string, string?> { ["A1"] = "same" }));
        var fileB = CreateWorkbook(
            new TestSheet("Sheet1", values: new Dictionary<string, string?> { ["A1"] = "same" }),
            new TestSheet("Sheet2", values: new Dictionary<string, string?> { ["A1"] = "new" }));

        try
        {
            var result = await comparer.CompareAsync(fileA, fileB, CreateMinimalOptions(), progress: null, CancellationToken.None);

            var diff = Assert.Single(result.Diffs.Where(x => x.What == "Sheet" && x.Kind == DiffKind.Added));
            Assert.Equal("Sheet2", diff.Sheet);
            Assert.Equal("Missing", diff.Before);
            Assert.Equal("Present", diff.After);
        }
        finally
        {
            File.Delete(fileA);
            File.Delete(fileB);
        }
    }

    [Fact]
    public async Task CompareAsync_WhenCellValueChanges_ShouldReturnModifiedValueDiff()
    {
        var comparer = new ExcelComparer.Infrastracture.ExcelComparer();
        var fileA = CreateWorkbook(new TestSheet("Sheet1", values: new Dictionary<string, string?> { ["A1"] = "old" }));
        var fileB = CreateWorkbook(new TestSheet("Sheet1", values: new Dictionary<string, string?> { ["A1"] = "new" }));

        try
        {
            var result = await comparer.CompareAsync(fileA, fileB, CreateMinimalOptions(), progress: null, CancellationToken.None);

            var diff = Assert.Single(result.Diffs.Where(x => x.What == "Value" && x.Address == "A1"));
            Assert.Equal(DiffKind.Modified, diff.Kind);
            Assert.Equal("old", diff.Before);
            Assert.Equal("new", diff.After);
        }
        finally
        {
            File.Delete(fileA);
            File.Delete(fileB);
        }
    }

    [Fact]
    public async Task CompareAsync_WhenFormulaChanges_ShouldReturnModifiedFormulaDiff()
    {
        var comparer = new ExcelComparer.Infrastracture.ExcelComparer();
        var fileA = CreateWorkbook(new TestSheet("Sheet1", formulas: new Dictionary<string, string?> { ["B1"] = "A1" }));
        var fileB = CreateWorkbook(new TestSheet("Sheet1", formulas: new Dictionary<string, string?> { ["B1"] = "A1+1" }));

        try
        {
            var result = await comparer.CompareAsync(fileA, fileB, CreateMinimalOptions(), progress: null, CancellationToken.None);

            var diff = Assert.Single(result.Diffs.Where(x => x.What == "Formula" && x.Address == "B1"));
            Assert.Equal(DiffKind.Modified, diff.Kind);
            Assert.Equal("A1", diff.Before);
            Assert.Equal("A1+1", diff.After);
        }
        finally
        {
            File.Delete(fileA);
            File.Delete(fileB);
        }
    }

    [Fact]
    public async Task CompareAsync_WhenHiddenSheetsAreExcluded_ShouldSkipHiddenSheetCellDiffs()
    {
        var comparer = new ExcelComparer.Infrastracture.ExcelComparer();
        var fileA = CreateWorkbook(new TestSheet("Hidden", hidden: true, values: new Dictionary<string, string?> { ["A1"] = "old" }));
        var fileB = CreateWorkbook(new TestSheet("Hidden", hidden: true, values: new Dictionary<string, string?> { ["A1"] = "new" }));
        var options = CreateMinimalOptions();
        options.IncludeHiddenSheets = false;

        try
        {
            var result = await comparer.CompareAsync(fileA, fileB, options, progress: null, CancellationToken.None);

            Assert.Empty(result.Diffs);
        }
        finally
        {
            File.Delete(fileA);
            File.Delete(fileB);
        }
    }

    [Fact]
    public async Task CompareAsync_WhenCancellationRequested_ShouldThrowOperationCanceledException()
    {
        var comparer = new ExcelComparer.Infrastracture.ExcelComparer();
        var fileA = CreateWorkbook(new TestSheet("Sheet1", values: new Dictionary<string, string?> { ["A1"] = "old" }));
        var fileB = CreateWorkbook(new TestSheet("Sheet1", values: new Dictionary<string, string?> { ["A1"] = "new" }));
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        try
        {
            await Assert.ThrowsAsync<OperationCanceledException>(async () =>
                await comparer.CompareAsync(fileA, fileB, CreateMinimalOptions(), progress: null, cts.Token));
        }
        finally
        {
            File.Delete(fileA);
            File.Delete(fileB);
        }
    }

    [Fact]
    public async Task CompareAsync_ShouldReportFinalProgress()
    {
        var comparer = new ExcelComparer.Infrastracture.ExcelComparer();
        var fileA = CreateWorkbook(new TestSheet("Sheet1", values: new Dictionary<string, string?> { ["A1"] = "same" }));
        var fileB = CreateWorkbook(new TestSheet("Sheet1", values: new Dictionary<string, string?> { ["A1"] = "same" }));
        var updates = new List<ProgressInfo>();
        var progress = new InlineProgress<ProgressInfo>(updates.Add);

        try
        {
            await comparer.CompareAsync(fileA, fileB, CreateMinimalOptions(), progress, CancellationToken.None);

            var finalUpdate = Assert.Single(updates.Where(x => x.Percent == 100));
            Assert.Equal("Finalizado.", finalUpdate.Message);
        }
        finally
        {
            File.Delete(fileA);
            File.Delete(fileB);
        }
    }

    private static ComparisonOptions CreateMinimalOptions()
        => new()
        {
            CompareValues = true,
            CompareFormulas = true,
            IncludeHiddenSheets = true,
            CompareSheetOrder = false,
            CompareWorkbookProperties = false,
            CompareUsedRange = false,
            CompareValidations = false,
            CompareConditionalFormats = false,
            CompareHiddenRowsCols = false,
            CompareCellFormat = false
        };

    private sealed class InlineProgress<T>(Action<T> onReport) : IProgress<T>
    {
        public void Report(T value)
        {
            onReport(value);
        }
    }
}
