using System.Windows.Controls;
using Microsoft.Extensions.DependencyInjection;
using OdooAutoCAD.App.Services;

namespace OdooAutoCAD.App.Views.Pages;

public partial class LogsPage : Page
{
    public LogsPage()
    {
        InitializeComponent();

        if (App.Services != null)
        {
            var logService = App.Services.GetRequiredService<IAppLogService>();
            FullLogPanel.DataContext = logService;
        }
    }
}
