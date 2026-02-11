using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Extensions.DependencyInjection;
using OdooAutoCAD.App.ViewModels;

namespace OdooAutoCAD.App.Views.Pages;

public partial class PRPage : Page
{
    public PRPage()
    {
        try
        {
            InitializeComponent();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"PRPage XAML init failed: {ex}");
            Content = new Border
            {
                Margin = new Thickness(24),
                Child = new StackPanel
                {
                    Children =
                    {
                        new TextBlock
                        {
                            Text = "Purchase Requisition",
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
                DataContext = App.Services.GetRequiredService<PurchaseRequisitionViewModel>();
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"PRPage ViewModel init failed: {ex}");
        }
    }
}
