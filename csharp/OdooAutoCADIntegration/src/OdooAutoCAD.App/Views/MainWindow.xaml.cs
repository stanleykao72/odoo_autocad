// OdooAutoCAD.App/Views/MainWindow.xaml.cs
// Main window code-behind

using System.Windows;
using System.Windows.Controls;
using Microsoft.Extensions.DependencyInjection;
using OdooAutoCAD.App.Services;
using OdooAutoCAD.App.ViewModels;

namespace OdooAutoCAD.App.Views;

/// <summary>
/// Main application window with side navigation.
/// </summary>
public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;
    private readonly INavigationService _navigationService;

    public MainWindow()
    {
        InitializeComponent();

        _viewModel = App.Services.GetRequiredService<MainViewModel>();
        _navigationService = App.Services.GetRequiredService<INavigationService>();

        // Wire the Frame to the navigation service before first navigation
        _navigationService.Frame = MainFrame;

        DataContext = _viewModel;

        // Navigate to dashboard by default
        _viewModel.NavigateCommand.Execute("Dashboard");
    }

    private void NavButton_Click(object sender, RoutedEventArgs e)
    {
        if (sender is Button button)
        {
            var pageName = button.Name.Replace("Btn", "");
            _viewModel.NavigateCommand.Execute(pageName);
        }
    }
}
