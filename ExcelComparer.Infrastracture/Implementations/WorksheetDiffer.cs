using DocumentFormat.OpenXml.Spreadsheet;
using ExcelComparer.Application.Models;
using ExcelComparer.Domain.Entities;
using ExcelComparer.Infrastructure.Interfaces;
using System.Text;

namespace ExcelComparer.Infrastructure.Implementations;

internal sealed class WorksheetDiffer : IWorksheetDiffer
{
    private const string EmptyAddress = "";
    private const string CellItem = "Cell";

    public ValueTask DiffUsedRangeAsync(Worksheet worksheetA, Worksheet worksheetB, string sheetName, ComparisonResult result)
    {
        var usedRangeA = GetUsedRange(worksheetA);
        var usedRangeB = GetUsedRange(worksheetB);

        if (!StringEquals(usedRangeA, usedRangeB))
        {
            result.Diffs.Add(new DiffItem(sheetName, EmptyAddress, DiffKind.Modified, "UsedRange", usedRangeA, usedRangeB));
        }

        return ValueTask.CompletedTask;
    }

    public void DiffDataValidations(Worksheet worksheetA, Worksheet worksheetB, string sheetName, ComparisonResult result)
    {
        var validationsA = GetDataValidationSummary(worksheetA);
        var validationsB = GetDataValidationSummary(worksheetB);

        if (!StringEquals(validationsA, validationsB))
        {
            result.Diffs.Add(new DiffItem(sheetName, EmptyAddress, DiffKind.Modified, "DataValidation", validationsA, validationsB));
        }
    }

    public ValueTask DiffConditionalFormattingAsync(Worksheet worksheetA, Worksheet worksheetB, string sheetName, ComparisonResult result)
    {
        var formattingA = worksheetA.Descendants<ConditionalFormatting>().Select(cf => cf.InnerText).ToList();
        var formattingB = worksheetB.Descendants<ConditionalFormatting>().Select(cf => cf.InnerText).ToList();

        if (!formattingA.SequenceEqual(formattingB))
        {
            result.Diffs.Add(new DiffItem(
                sheetName,
                EmptyAddress,
                DiffKind.Modified,
                "ConditionalFormatting",
                string.Join("|", formattingA),
                string.Join("|", formattingB)));
        }

        return ValueTask.CompletedTask;
    }

    public ValueTask DiffHiddenRowsColsAsync(Worksheet worksheetA, Worksheet worksheetB, string sheetName, ComparisonResult result)
    {
        var hiddenColumnsA = GetHiddenColumns(worksheetA);
        var hiddenColumnsB = GetHiddenColumns(worksheetB);
        if (!hiddenColumnsA.SequenceEqual(hiddenColumnsB))
        {
            result.Diffs.Add(new DiffItem(sheetName, EmptyAddress, DiffKind.Modified, "HiddenColumns", string.Join(",", hiddenColumnsA), string.Join(",", hiddenColumnsB)));
        }

        var hiddenRowsA = GetHiddenRows(worksheetA);
        var hiddenRowsB = GetHiddenRows(worksheetB);
        if (!hiddenRowsA.SequenceEqual(hiddenRowsB))
        {
            result.Diffs.Add(new DiffItem(sheetName, EmptyAddress, DiffKind.Modified, "HiddenRows", string.Join(",", hiddenRowsA), string.Join(",", hiddenRowsB)));
        }

        return ValueTask.CompletedTask;
    }

    public ValueTask DiffCellsAsync(
        string sheetName,
        Dictionary<string, CellInfo> cellsA,
        Dictionary<string, CellInfo> cellsB,
        ComparisonResult result,
        ComparisonOptions options)
    {
        var seenAddresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var diffs = new List<DiffItem>(cellsA.Count + cellsB.Count);

        foreach (var address in cellsA.Keys)
        {
            ProcessCell(address);
        }

        foreach (var address in cellsB.Keys)
        {
            ProcessCell(address);
        }

        if (diffs.Count > 0)
        {
            result.Diffs.AddRange(diffs);
        }

        return ValueTask.CompletedTask;

        void ProcessCell(string address)
        {
            if (!seenAddresses.Add(address))
            {
                return;
            }

            var hasA = cellsA.TryGetValue(address, out var cellA);
            var hasB = cellsB.TryGetValue(address, out var cellB);

            if (hasA && !hasB)
            {
                diffs.Add(new DiffItem(sheetName, address, DiffKind.Removed, CellItem, Summarize(cellA, options), null));
                return;
            }

            if (!hasA && hasB)
            {
                diffs.Add(new DiffItem(sheetName, address, DiffKind.Added, CellItem, null, Summarize(cellB, options)));
                return;
            }

            AddCellChanges(diffs, sheetName, address, cellA!, cellB!, options);
        }
    }

    private static void AddCellChanges(List<DiffItem> diffs, string sheetName, string address, CellInfo before, CellInfo after, ComparisonOptions options)
    {
        AddModifiedDiffIfNeeded(diffs, sheetName, address, "Value", before.ValueText, after.ValueText, options.CompareValues);
        AddModifiedDiffIfNeeded(diffs, sheetName, address, "Formula", before.FormulaText, after.FormulaText, options.CompareFormulas);

        if (!options.CompareCellFormat)
        {
            return;
        }

        AddModifiedDiffIfNeeded(diffs, sheetName, address, "StyleIndex", before.StyleIndex?.ToString(), after.StyleIndex?.ToString(), true);
        AddModifiedDiffIfNeeded(diffs, sheetName, address, "NumberFormat", before.NumberFormatCode, after.NumberFormatCode, true);
    }

    private static void AddModifiedDiffIfNeeded(List<DiffItem> diffs, string sheetName, string address, string what, string? before, string? after, bool shouldCompare)
    {
        if (!shouldCompare || StringEquals(before, after))
        {
            return;
        }

        diffs.Add(new DiffItem(sheetName, address, DiffKind.Modified, what, before, after));
    }

    private static List<string> GetHiddenColumns(Worksheet worksheet)
        => worksheet.Elements<Columns>().FirstOrDefault()?
            .Elements<Column>()
            .Where(c => c.Hidden != null && c.Hidden.Value)
            .Select(c => $"{c.Min}-{c.Max}")
            .OrderBy(x => x)
            .ToList()
            ?? [];

    private static List<uint> GetHiddenRows(Worksheet worksheet)
        => worksheet.Descendants<Row>()
            .Where(r => r.Hidden != null && r.Hidden.Value)
            .Select(r => r.RowIndex!.Value)
            .OrderBy(x => x)
            .ToList();

    private static string GetDataValidationSummary(Worksheet worksheet)
        => string.Join("|", worksheet.Descendants<DataValidation>().Select(v => v.OuterXml));

    private static string GetUsedRange(Worksheet worksheet)
    {
        var reference = worksheet.SheetDimension?.Reference?.Value;
        if (!string.IsNullOrWhiteSpace(reference))
        {
            return reference;
        }

        var cells = worksheet.Descendants<Cell>()
            .Select(c => c.CellReference?.Value)
            .Where(c => !string.IsNullOrWhiteSpace(c))
            .Cast<string>()
            .ToList();

        if (cells.Count == 0)
        {
            return string.Empty;
        }

        var minRow = int.MaxValue;
        var maxRow = 0;
        var minCol = int.MaxValue;
        var maxCol = 0;

        foreach (var cell in cells)
        {
            var (row, column) = ParseCellReference(cell);
            if (row <= 0 || column <= 0)
            {
                continue;
            }

            minRow = Math.Min(minRow, row);
            maxRow = Math.Max(maxRow, row);
            minCol = Math.Min(minCol, column);
            maxCol = Math.Max(maxCol, column);
        }

        if (maxRow == 0 || maxCol == 0)
        {
            return string.Empty;
        }

        return $"{ToColumnName(minCol)}{minRow}:{ToColumnName(maxCol)}{maxRow}";
    }

    private static (int Row, int Column) ParseCellReference(string cellReference)
    {
        var splitIndex = 0;
        while (splitIndex < cellReference.Length && char.IsLetter(cellReference[splitIndex]))
        {
            splitIndex++;
        }

        var column = ToColumnIndex(cellReference.Substring(0, splitIndex));
        var row = int.TryParse(cellReference.Substring(splitIndex), out var parsedRow) ? parsedRow : 0;
        return (row, column);
    }

    private static int ToColumnIndex(string name)
    {
        var result = 0;
        foreach (var ch in name.ToUpperInvariant())
        {
            result = (result * 26) + (ch - 'A' + 1);
        }

        return result;
    }

    private static string ToColumnName(int index)
    {
        var builder = new StringBuilder();
        while (index > 0)
        {
            var remainder = (index - 1) % 26;
            builder.Insert(0, (char)('A' + remainder));
            index = (index - 1) / 26;
        }

        return builder.ToString();
    }

    private static string? Summarize(CellInfo? cell, ComparisonOptions options)
    {
        if (cell is null)
        {
            return null;
        }

        var builder = new StringBuilder();
        if (options.CompareFormulas && !string.IsNullOrWhiteSpace(cell.FormulaText))
        {
            builder.Append('=').Append(cell.FormulaText);
        }

        if (options.CompareValues && !string.IsNullOrWhiteSpace(cell.ValueText))
        {
            if (builder.Length > 0)
            {
                builder.Append(" | ");
            }

            builder.Append(cell.ValueText);
        }

        if (options.CompareCellFormat)
        {
            if (builder.Length > 0)
            {
                builder.Append(" | ");
            }

            builder.Append("style:")
                .Append(cell.StyleIndex?.ToString() ?? string.Empty)
                .Append(" nf:")
                .Append(cell.NumberFormatCode ?? string.Empty);
        }

        return builder.Length == 0 ? null : builder.ToString();
    }

    private static bool StringEquals(string? left, string? right)
        => string.Equals(left, right, StringComparison.Ordinal);
}
