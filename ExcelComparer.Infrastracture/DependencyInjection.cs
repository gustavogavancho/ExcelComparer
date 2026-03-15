using ExcelComparer.Application.Contracts;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelComparer.Infrastracture;

public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructure(this IServiceCollection services)
    {
        services.AddSingleton<IExcelComparer, ExcelComparer>();
        return services;
    }
}
