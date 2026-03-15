using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelComparer.Infrastructure.UnitTests;

internal static class TestWorkbookFactory
{
    public static string CreateWorkbook(params TestSheet[] sheets)
    {
        var path = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.xlsx");

        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var sheetsElement = workbookPart.Workbook.AppendChild(new Sheets());
        uint sheetId = 1;

        foreach (var sheet in sheets)
        {
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            var worksheet = new Worksheet(sheetData);

            foreach (var rowGroup in GetCellDefinitions(sheet).GroupBy(x => x.RowIndex).OrderBy(x => x.Key))
            {
                var row = new Row { RowIndex = (uint)rowGroup.Key };
                foreach (var cellDefinition in rowGroup.OrderBy(x => x.ColumnIndex))
                {
                    row.Append(CreateCell(cellDefinition));
                }

                sheetData.Append(row);
            }

            worksheetPart.Worksheet = worksheet;
            worksheetPart.Worksheet.Save();

            var workbookSheet = new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId++,
                Name = sheet.Name,
                State = sheet.Hidden ? SheetStateValues.Hidden : SheetStateValues.Visible
            };

            sheetsElement.Append(workbookSheet);
        }

        workbookPart.Workbook.Save();
        return path;
    }

    private static IEnumerable<TestCell> GetCellDefinitions(TestSheet sheet)
    {
        var addresses = sheet.Values.Keys
            .Concat(sheet.Formulas.Keys)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (var address in addresses)
        {
            ParseAddress(address, out var rowIndex, out var columnIndex);
            sheet.Values.TryGetValue(address, out var value);
            sheet.Formulas.TryGetValue(address, out var formula);

            yield return new TestCell(address, rowIndex, columnIndex, value, formula);
        }
    }

    private static Cell CreateCell(TestCell definition)
    {
        var cell = new Cell { CellReference = definition.Address };

        if (!string.IsNullOrWhiteSpace(definition.Formula))
        {
            cell.CellFormula = new CellFormula(definition.Formula);
        }

        if (definition.Value is not null)
        {
            cell.DataType = CellValues.InlineString;
            cell.InlineString = new InlineString(new Text(definition.Value) { Space = SpaceProcessingModeValues.Preserve });
        }

        return cell;
    }

    private static void ParseAddress(string address, out int rowIndex, out int columnIndex)
    {
        var splitIndex = 0;
        while (splitIndex < address.Length && char.IsLetter(address[splitIndex]))
        {
            splitIndex++;
        }

        rowIndex = int.Parse(address[splitIndex..]);
        columnIndex = 0;

        foreach (var ch in address[..splitIndex].ToUpperInvariant())
        {
            columnIndex = (columnIndex * 26) + (ch - 'A' + 1);
        }
    }

    internal sealed class TestSheet
    {
        public TestSheet(
            string name,
            bool hidden = false,
            IReadOnlyDictionary<string, string?>? values = null,
            IReadOnlyDictionary<string, string?>? formulas = null)
        {
            Name = name;
            Hidden = hidden;
            Values = values ?? new Dictionary<string, string?>();
            Formulas = formulas ?? new Dictionary<string, string?>();
        }

        public string Name { get; }

        public bool Hidden { get; }

        public IReadOnlyDictionary<string, string?> Values { get; }

        public IReadOnlyDictionary<string, string?> Formulas { get; }
    }

    private sealed record TestCell(string Address, int RowIndex, int ColumnIndex, string? Value, string? Formula);
}
