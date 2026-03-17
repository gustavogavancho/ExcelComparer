using ExcelComparer.Application.Interfaces;
using ExcelComparer.Infrastructure.Implementations;
using ExcelComparer.Infrastructure.Interfaces;
using Microsoft.Extensions.DependencyInjection;
using OpenXmlExcelComparer = ExcelComparer.Infrastructure.Implementations.ExcelComparer;

namespace ExcelComparer.Infrastructure;

public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructure(this IServiceCollection services)
    {
        services.AddSingleton<IOpenXmlWorkbookReader, OpenXmlWorkbookReader>();
        services.AddSingleton<IWorkbookDiffer, WorkbookDiffer>();
        services.AddSingleton<IWorksheetDiffer, WorksheetDiffer>();
        services.AddSingleton<IExcelComparer>(serviceProvider => new OpenXmlExcelComparer(
            serviceProvider.GetRequiredService<IOpenXmlWorkbookReader>(),
            serviceProvider.GetRequiredService<IWorkbookDiffer>(),
            serviceProvider.GetRequiredService<IWorksheetDiffer>()));
        return services;
    }
}
