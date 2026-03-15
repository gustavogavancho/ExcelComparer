namespace ExcelComparer.Application.Models;

public sealed record DiffItem(
    string Sheet,
    string Address,
    DiffKind Kind,
    string What,
    string? Before,
    string? After
);
