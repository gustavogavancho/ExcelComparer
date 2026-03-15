using ExcelComparer.Domain.Entities;

namespace ExcelComparer.Domain.UnitTests;

public class DomainEntitiesTests
{
    [Fact]
    public void WorkbookInfo_ShouldUseCaseInsensitiveSheetDictionary()
    {
        var workbook = new WorkbookInfo();
        var sheet = new SheetInfo
        {
            Name = "Summary",
            SheetId = "1",
            RelId = "rId1",
            Hidden = false,
            VeryHidden = false
        };

        workbook.SheetsByName["Summary"] = sheet;

        Assert.True(workbook.SheetsByName.ContainsKey("summary"));
        Assert.Same(sheet, workbook.SheetsByName["SUMMARY"]);
    }

    [Fact]
    public void SheetInfo_And_CellInfo_ShouldStoreProvidedValues()
    {
        var sheet = new SheetInfo
        {
            Name = "HiddenSheet",
            SheetId = "2",
            RelId = "rId2",
            Hidden = true,
            VeryHidden = false
        };

        var cell = new CellInfo
        {
            ValueText = "42",
            FormulaText = "SUM(A1:A2)",
            StyleIndex = 3,
            NumberFormatCode = "0.00"
        };

        Assert.Equal("HiddenSheet", sheet.Name);
        Assert.Equal("2", sheet.SheetId);
        Assert.Equal("rId2", sheet.RelId);
        Assert.True(sheet.Hidden);
        Assert.False(sheet.VeryHidden);

        Assert.Equal("42", cell.ValueText);
        Assert.Equal("SUM(A1:A2)", cell.FormulaText);
        Assert.Equal((uint)3, cell.StyleIndex);
        Assert.Equal("0.00", cell.NumberFormatCode);
    }
}
