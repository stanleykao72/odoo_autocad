using System.ComponentModel;
using System.Windows.Controls;
using Microsoft.Extensions.DependencyInjection;
using OdooAutoCAD.App.ViewModels;

namespace OdooAutoCAD.App.Views.Pages;

public partial class SettingsPage : Page
{
    private readonly SettingsViewModel _viewModel;
    private bool _syncingToken;

    public SettingsPage()
    {
        InitializeComponent();

        _viewModel = App.Services.GetRequiredService<SettingsViewModel>();
        DataContext = _viewModel;

        // Wire PasswordBox ↔ ViewModel (PasswordBox.Password is not a DependencyProperty)
        TokenPasswordBox.PasswordChanged += OnPasswordBoxChanged;
        _viewModel.PropertyChanged += OnViewModelPropertyChanged;

        // Load settings when page is ready
        Loaded += async (_, _) => await _viewModel.LoadSettingsCommand.ExecuteAsync(null);
    }

    private void OnPasswordBoxChanged(object sender, System.Windows.RoutedEventArgs e)
    {
        if (_syncingToken) return;
        _syncingToken = true;
        _viewModel.OdooApiToken = TokenPasswordBox.Password;
        _syncingToken = false;
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(SettingsViewModel.OdooApiToken))
        {
            if (_syncingToken) return;
            _syncingToken = true;
            TokenPasswordBox.Password = _viewModel.OdooApiToken;
            _syncingToken = false;
        }
        else if (e.PropertyName == nameof(SettingsViewModel.IsTokenVisible))
        {
            ToggleTokenButton.Content = _viewModel.IsTokenVisible ? "Hide" : "Show";
        }
    }
}
