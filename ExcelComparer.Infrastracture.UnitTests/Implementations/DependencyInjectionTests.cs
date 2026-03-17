using ExcelComparer.Application.Interfaces;
using ExcelComparer.Infrastructure;
using ExcelComparer.Infrastructure.Implementations;
using ExcelComparer.Infrastructure.Interfaces;
using Microsoft.Extensions.DependencyInjection;
using OpenXmlExcelComparer = ExcelComparer.Infrastructure.Implementations.ExcelComparer;

namespace ExcelComparer.Infrastructure.UnitTests;

public class DependencyInjectionTests
{
    [Fact]
    public void AddInfrastructure_ShouldRegisterInfrastructureServices()
    {
        var services = new ServiceCollection();

        services.AddInfrastructure();

        AssertRegistration<IOpenXmlWorkbookReader, OpenXmlWorkbookReader>(services);
        AssertRegistration<IWorkbookDiffer, WorkbookDiffer>(services);
        AssertRegistration<IWorksheetDiffer, WorksheetDiffer>(services);

        var comparerRegistration = Assert.Single(services.Where(x => x.ServiceType == typeof(IExcelComparer)));
        Assert.Equal(ServiceLifetime.Singleton, comparerRegistration.Lifetime);
        Assert.Null(comparerRegistration.ImplementationType);
        Assert.NotNull(comparerRegistration.ImplementationFactory);
    }

    [Fact]
    public void AddInfrastructure_ComparerFactory_ShouldCreateComparerUsingRegisteredCollaborators()
    {
        var services = new ServiceCollection();

        services.AddInfrastructure();

        var comparerRegistration = Assert.Single(services.Where(x => x.ServiceType == typeof(IExcelComparer)));
        var serviceProvider = new TestServiceProvider(new Dictionary<Type, object>
        {
            [typeof(IOpenXmlWorkbookReader)] = new OpenXmlWorkbookReader(),
            [typeof(IWorkbookDiffer)] = new WorkbookDiffer(),
            [typeof(IWorksheetDiffer)] = new WorksheetDiffer()
        });

        var service = comparerRegistration.ImplementationFactory!(serviceProvider);

        Assert.NotNull(service);
        Assert.IsType<OpenXmlExcelComparer>(service);
    }

    private static void AssertRegistration<TService, TImplementation>(ServiceCollection services)
        where TService : class
        where TImplementation : class, TService
    {
        var registration = Assert.Single(services.Where(x => x.ServiceType == typeof(TService)));
        Assert.Equal(ServiceLifetime.Singleton, registration.Lifetime);
        Assert.Equal(typeof(TImplementation), registration.ImplementationType);
    }

    private sealed class TestServiceProvider(IReadOnlyDictionary<Type, object> services) : IServiceProvider
    {
        public object? GetService(Type serviceType)
            => services.TryGetValue(serviceType, out var service) ? service : null;
    }
}
