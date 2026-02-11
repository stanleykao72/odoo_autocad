using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Extensions.DependencyInjection;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Core.AutoCAD;

namespace OdooAutoCAD.App.Views.Pages;

public partial class AutoCADPage : Page
{
    public AutoCADPage()
    {
        try
        {
            InitializeComponent();
        }
        catch (Exception ex)
        {
            // XAML parse failed — create minimal fallback content
            Debug.WriteLine($"AutoCADPage XAML init failed: {ex}");
            Content = new Border
            {
                Margin = new Thickness(40),
                Child = new StackPanel
                {
                    Children =
                    {
                        new TextBlock
                        {
                            Text = "AutoCAD Integration",
                            FontSize = 28,
                            FontWeight = FontWeights.Bold,
                            Margin = new Thickness(0, 0, 0, 16)
                        },
                        new TextBlock
                        {
                            Text = $"Page failed to load: {ex.Message}",
                            FontSize = 14,
                            Foreground = Brushes.Red,
                            TextWrapping = TextWrapping.Wrap,
                            Margin = new Thickness(0, 0, 0, 8)
                        },
                        new TextBlock
                        {
                            Text = ex.InnerException?.Message ?? string.Empty,
                            FontSize = 12,
                            Foreground = Brushes.Gray,
                            TextWrapping = TextWrapping.Wrap
                        }
                    }
                }
            };
            return;
        }

        try
        {
            var viewModel = App.Services.GetRequiredService<AutoCADViewModel>();
            DataContext = viewModel;
        }
        catch (Exception ex)
        {
            // ViewModel DI failed — page still renders its XAML content
            Debug.WriteLine($"AutoCADPage ViewModel init failed: {ex}");
        }
    }

    private void ListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        try
        {
            if (sender is ListViewItem item && item.DataContext is LayoutInfo layout)
            {
                if (DataContext is AutoCADViewModel viewModel)
                {
                    viewModel.SelectLayoutCommand.Execute(layout);
                }
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Layout selection error: {ex.Message}");
        }
    }
}
