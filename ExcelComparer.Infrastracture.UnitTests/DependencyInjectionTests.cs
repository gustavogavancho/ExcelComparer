using ExcelComparer.Application.Contracts;
using ExcelComparer.Infrastracture;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelComparer.Infrastracture.UnitTests;

public class DependencyInjectionTests
{
    [Fact]
    public void AddInfrastructure_ShouldRegisterExcelComparerService()
    {
        var services = new ServiceCollection();

        services.AddInfrastructure();

        var registration = Assert.Single(services.Where(x => x.ServiceType == typeof(IExcelComparer)));

        Assert.Equal(ServiceLifetime.Singleton, registration.Lifetime);
        Assert.Equal(typeof(ExcelComparer.Infrastracture.ExcelComparer), registration.ImplementationType);
    }
}
