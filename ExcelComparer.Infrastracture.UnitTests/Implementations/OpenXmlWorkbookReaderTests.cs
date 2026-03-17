using DocumentFormat.OpenXml.Packaging;
using ExcelComparer.Infrastructure.Implementations;
using static ExcelComparer.Infrastructure.UnitTests.TestWorkbookFactory;

namespace ExcelComparer.Infrastructure.UnitTests;

public class OpenXmlWorkbookReaderTests
{
    [Fact]
    public void ReadWorkbook_ShouldPopulateSheetMetadata()
    {
        var reader = new OpenXmlWorkbookReader();
        var file = CreateWorkbook(
            new TestSheet("VisibleSheet"),
            new TestSheet("HiddenSheet", hidden: true));

        try
        {
            using var document = SpreadsheetDocument.Open(file, false);

            var workbook = reader.ReadWorkbook(document);

            Assert.Equal(new[] { "VisibleSheet", "HiddenSheet" }, workbook.SheetOrder);
            Assert.False(workbook.SheetsByName["VisibleSheet"].Hidden);
            Assert.True(workbook.SheetsByName["HiddenSheet"].Hidden);
            Assert.Equal("HiddenSheet", workbook.SheetsByName["HiddenSheet"].Name);
        }
        finally
        {
            File.Delete(file);
        }
    }
}
