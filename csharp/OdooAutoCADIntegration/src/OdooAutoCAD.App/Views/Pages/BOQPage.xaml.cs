using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Extensions.DependencyInjection;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.App.ViewModels.Dialogs;
using OdooAutoCAD.App.Views.Dialogs;
using OdooAutoCAD.Core.BOQ;
using OdooAutoCAD.Core.Odoo;

namespace OdooAutoCAD.App.Views.Pages;

public partial class BOQPage : Page
{
    public BOQPage()
    {
        try
        {
            InitializeComponent();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"BOQPage XAML init failed: {ex}");
            Content = new Border
            {
                Margin = new Thickness(24),
                Child = new StackPanel
                {
                    Children =
                    {
                        new TextBlock
                        {
                            Text = "BOQ Manager",
                            FontSize = 28,
                            FontWeight = FontWeights.Bold,
                            Margin = new Thickness(0, 0, 0, 16)
                        },
                        new TextBlock
                        {
                            Text = $"Page failed to load: {ex.Message}",
                            FontSize = 14,
                            Foreground = Brushes.Red,
                            TextWrapping = TextWrapping.Wrap
                        }
                    }
                }
            };
            return;
        }

        try
        {
            if (App.Services != null)
            {
                DataContext = App.Services.GetRequiredService<BOQViewModel>();
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"BOQPage ViewModel init failed: {ex}");
        }
    }

    private void ManageMappingsBtn_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (App.Services == null) return;

            var odooService = App.Services.GetRequiredService<IOdooService>();
            var boqProcessor = App.Services.GetRequiredService<IBOQProcessor>();

            var vm = new ProductMappingDialogViewModel(odooService, boqProcessor)
            {
                IsManagementMode = true
            };
            vm.LoadMappingsCommand.Execute(null);

            var dialog = new ProductMappingDialog(vm)
            {
                Owner = Window.GetWindow(this)
            };
            dialog.ShowDialog();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Manage Mappings dialog error: {ex}");
        }
    }
}
