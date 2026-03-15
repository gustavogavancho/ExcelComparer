using ExcelComparer.Application.Models;

namespace ExcelComparer.Application.Contracts;

public interface IExcelComparer
{
    ValueTask<ComparisonResult> CompareAsync(string fileA, string fileB, ComparisonOptions options, IProgress<ProgressInfo>? progress, CancellationToken ct);
}
