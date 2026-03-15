using ExcelComparer.Infrastructure;
using ExcelComparer.WPFUI.ViewModels;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System.Windows;

namespace ExcelComparer.WPFUI;

public partial class App : System.Windows.Application
{
    private readonly IHost _host;

    public App()
    {
        System.Windows.FrameworkCompatibilityPreferences.KeepTextBoxDisplaySynchronizedWithTextProperty = false;
        _host = CreateHostBuilder().Build();
    }

    public static IHostBuilder CreateHostBuilder(string[] args = null)
    {
        return Host.CreateDefaultBuilder(args)
            .ConfigureServices(services =>
            {
                services.AddInfrastructure();
                services.AddSingleton<MainViewModel>();
                services.AddSingleton<MainWindow>();
            });
    }

    private async void Application_Startup(object sender, StartupEventArgs e)
    {
        await _host.StartAsync();

        Window window = _host.Services.GetRequiredService<MainWindow>();
        window.Show();
    }

    protected override async void OnExit(ExitEventArgs e)
    {
        await _host.StopAsync();
        _host.Dispose();
        base.OnExit(e);
    }
}