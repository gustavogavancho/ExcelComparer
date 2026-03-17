using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelComparer.Application.Models;
using ExcelComparer.Domain.Entities;
using ExcelComparer.Infrastructure.Interfaces;
using System.Globalization;

namespace ExcelComparer.Infrastructure.Implementations;

internal sealed class OpenXmlWorkbookReader : IOpenXmlWorkbookReader
{
    public ValueTask<Dictionary<int, string?>> BuildNumberFormatCacheAsync(WorkbookPart workbookPart)
    {
        var result = new Dictionary<int, string?>();
        var styles = workbookPart.WorkbookStylesPart?.Stylesheet;
        if (styles is null)
        {
            return new ValueTask<Dictionary<int, string?>>(result);
        }

        var cellFormats = styles.CellFormats?.Elements<CellFormat>().ToList();
        if (cellFormats is null)
        {
            return new ValueTask<Dictionary<int, string?>>(result);
        }

        var numberingFormats = styles.NumberingFormats?
            .Elements<NumberingFormat>()
            .Where(n => n.NumberFormatId != null)
            .ToDictionary(n => (int)n.NumberFormatId!.Value, n => n.FormatCode?.Value)
            ?? new Dictionary<int, string?>();

        for (int i = 0; i < cellFormats.Count; i++)
        {
            var cellFormat = cellFormats[i];
            if (cellFormat.NumberFormatId == null)
            {
                result[i] = null;
                continue;
            }

            var numberFormatId = (int)cellFormat.NumberFormatId.Value;
            result[i] = numberingFormats.TryGetValue(numberFormatId, out var code)
                ? code
                : numberFormatId.ToString(CultureInfo.InvariantCulture);
        }

        return new ValueTask<Dictionary<int, string?>>(result);
    }

    public WorkbookInfo ReadWorkbook(SpreadsheetDocument document)
    {
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart no encontrado.");
        var workbook = workbookPart.Workbook ?? throw new InvalidOperationException("Workbook no encontrado.");

        var info = new WorkbookInfo();

        foreach (var sheet in workbook.Sheets!.OfType<Sheet>())
        {
            var name = sheet.Name?.Value ?? "(sin nombre)";
            var state = sheet.State?.Value;

            info.SheetsByName[name] = new SheetInfo
            {
                Name = name,
                SheetId = sheet.SheetId?.Value.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                RelId = sheet.Id?.Value ?? string.Empty,
                Hidden = state == SheetStateValues.Hidden,
                VeryHidden = state == SheetStateValues.VeryHidden
            };

            info.SheetOrder.Add(name);
        }

        return info;
    }

    public async ValueTask<Dictionary<string, CellInfo>> ReadCellsAsync(
        WorkbookPart workbookPart,
        Worksheet worksheet,
        SharedStringTable? sharedStringTable,
        ComparisonOptions options,
        Dictionary<int, string?>? numberFormatCache)
    {
        var cells = new Dictionary<string, CellInfo>(StringComparer.OrdinalIgnoreCase);
        var sheetData = worksheet.Elements<SheetData>().FirstOrDefault();
        if (sheetData is null)
        {
            return cells;
        }

        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                var address = cell.CellReference?.Value;
                if (string.IsNullOrWhiteSpace(address))
                {
                    continue;
                }

                var formula = cell.CellFormula?.Text;
                var value = await ReadCellValueAsTextAsync(cell, sharedStringTable);
                var styleIndex = GetStyleIndex(cell, options);
                var numberFormatCode = ResolveNumberFormatCode(workbookPart, styleIndex, numberFormatCache, options);

                if (value is null && formula is null && !options.CompareCellFormat)
                {
                    continue;
                }

                cells[address] = new CellInfo
                {
                    ValueText = value,
                    FormulaText = formula,
                    StyleIndex = styleIndex,
                    NumberFormatCode = numberFormatCode
                };
            }
        }

        return cells;
    }

    private static uint? GetStyleIndex(Cell cell, ComparisonOptions options)
        => options.CompareCellFormat ? cell.StyleIndex?.Value : null;

    private static string? ResolveNumberFormatCode(
        WorkbookPart workbookPart,
        uint? styleIndex,
        Dictionary<int, string?>? numberFormatCache,
        ComparisonOptions options)
    {
        if (!options.CompareCellFormat || styleIndex is null)
        {
            return null;
        }

        if (numberFormatCache != null && numberFormatCache.TryGetValue((int)styleIndex.Value, out var cachedCode))
        {
            return cachedCode;
        }

        return ResolveNumberFormatCode(workbookPart, styleIndex);
    }

    private static ValueTask<string?> ReadCellValueAsTextAsync(Cell cell, SharedStringTable? sharedStringTable)
    {
        if (cell.CellValue is null)
        {
            if (cell.DataType?.Value == CellValues.InlineString)
            {
                return new ValueTask<string?>(cell.InlineString?.Text?.Text ?? cell.InlineString?.InnerText);
            }

            return new ValueTask<string?>((string?)null);
        }

        var raw = cell.CellValue.Text;
        if (cell.DataType?.Value != CellValues.SharedString)
        {
            return new ValueTask<string?>(raw);
        }

        if (sharedStringTable is null || !int.TryParse(raw, out var index))
        {
            return new ValueTask<string?>(raw);
        }

        var item = sharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(index);
        return new ValueTask<string?>(item?.InnerText ?? raw);
    }

    private static string? ResolveNumberFormatCode(WorkbookPart workbookPart, uint? styleIndex)
    {
        if (styleIndex is null)
        {
            return null;
        }

        var styles = workbookPart.WorkbookStylesPart?.Stylesheet;
        if (styles is null)
        {
            return null;
        }

        var cellFormats = styles.CellFormats?.Elements<CellFormat>().ToList();
        if (cellFormats is null)
        {
            return null;
        }

        var index = (int)styleIndex.Value;
        if (index < 0 || index >= cellFormats.Count)
        {
            return null;
        }

        var cellFormat = cellFormats[index];
        if (cellFormat.NumberFormatId == null)
        {
            return null;
        }

        var numberFormatId = (int)cellFormat.NumberFormatId.Value;
        var customFormat = styles.NumberingFormats?
            .Elements<NumberingFormat>()
            .FirstOrDefault(n => n.NumberFormatId != null && n.NumberFormatId.Value == numberFormatId);

        return customFormat?.FormatCode?.Value ?? numberFormatId.ToString(CultureInfo.InvariantCulture);
    }
}
