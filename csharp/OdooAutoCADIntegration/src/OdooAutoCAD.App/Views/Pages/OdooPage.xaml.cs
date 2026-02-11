using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Extensions.DependencyInjection;
using OdooAutoCAD.App.ViewModels;

namespace OdooAutoCAD.App.Views.Pages;

public partial class OdooPage : Page
{
    public OdooPage()
    {
        try
        {
            InitializeComponent();
        }
        catch (Exception ex)
        {
            // XAML parse failed — create minimal fallback content
            Debug.WriteLine($"OdooPage XAML init failed: {ex}");
            Content = new Border
            {
                Margin = new Thickness(24),
                Child = new StackPanel
                {
                    Children =
                    {
                        new TextBlock
                        {
                            Text = "Odoo Connection",
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
            if (App.Services != null)
            {
                DataContext = App.Services.GetRequiredService<OdooConnectionViewModel>();
            }
        }
        catch (Exception ex)
        {
            // ViewModel DI failed — page still renders its XAML content
            Debug.WriteLine($"OdooPage ViewModel init failed: {ex}");
        }
    }
}
