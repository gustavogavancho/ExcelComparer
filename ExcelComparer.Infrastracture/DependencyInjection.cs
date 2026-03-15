using ExcelComparer.Application.Contracts;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelComparer.Infrastructure;

public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructure(this IServiceCollection services)
    {
        services.AddSingleton<IOpenXmlWorkbookReader, OpenXmlWorkbookReader>();
        services.AddSingleton<IWorkbookDiffer, WorkbookDiffer>();
        services.AddSingleton<IWorksheetDiffer, WorksheetDiffer>();
        services.AddSingleton<IExcelComparer>(serviceProvider => new ExcelComparer(
            serviceProvider.GetRequiredService<IOpenXmlWorkbookReader>(),
            serviceProvider.GetRequiredService<IWorkbookDiffer>(),
            serviceProvider.GetRequiredService<IWorksheetDiffer>()));
        return services;
    }
}
