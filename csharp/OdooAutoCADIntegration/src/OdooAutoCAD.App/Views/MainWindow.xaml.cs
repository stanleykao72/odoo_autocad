// OdooAutoCAD.App/Views/MainWindow.xaml.cs
// Main window code-behind

using System.Windows;
using System.Windows.Controls;
using Microsoft.Extensions.DependencyInjection;
using OdooAutoCAD.App.ViewModels;

namespace OdooAutoCAD.App.Views;

/// <summary>
/// Main application window with side navigation.
/// </summary>
public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;

    public MainWindow()
    {
        InitializeComponent();

        _viewModel = App.Services.GetRequiredService<MainViewModel>();
        DataContext = _viewModel;

        // Navigate to dashboard by default
        NavigateTo("Dashboard");
    }

    private void NavButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button button)
        {
            var pageName = button.Name.Replace("Btn", "");
            NavigateTo(pageName);
        }
    }

    private void NavigateTo(string pageName)
    {
        _viewModel.CurrentPageTitle = pageName switch
        {
            "Dashboard" => "Dashboard",
            "AutoCAD" => "AutoCAD Integration",
            "Odoo" => "Odoo Connection",
            "BOQ" => "BOQ Manager",
            "PR" => "Purchase Requisition",
            "MCP" => "AI Assistant (MCP)",
            "Settings" => "Settings",
            _ => pageName
        };

        // TODO: Navigate to actual pages
        // MainFrame.Navigate(new Uri($"Views/Pages/{pageName}Page.xaml", UriKind.Relative));
    }
}
