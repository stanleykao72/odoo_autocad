using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Media;
using Microsoft.Extensions.DependencyInjection;
using OdooAutoCAD.App.ViewModels;
using OdooAutoCAD.Core.Odoo;

namespace OdooAutoCAD.App.Views.Pages;

public partial class ParameterConfigPage : Page
{
    private bool _isUpdatingFromSelection;

    public ParameterConfigPage()
    {
        try
        {
            InitializeComponent();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"ParameterConfigPage XAML init failed: {ex}");
            Content = new Border
            {
                Margin = new Thickness(24),
                Child = new StackPanel
                {
                    Children =
                    {
                        new TextBlock
                        {
                            Text = "Parameter Configuration",
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
                var vm = App.Services.GetRequiredService<ParameterConfigViewModel>();
                DataContext = vm;

                SetupFilterableComboBox(CbProducts, "Name");
                SetupFilterableComboBox(CbSpecs, "Value");
                SetupFilterableComboBox(CbCategories, "Value");
                SetupFilterableComboBox(CbOperationFlows, "Value");
                SetupFilterableComboBox(CbSurfaceTreatments, "Value");
                SetupFilterableComboBox(CbColors, "Name");

                Loaded += async (s, e) =>
                {
                    if (vm.Products.Count == 0)
                    {
                        await vm.LoadOptionsCommand.ExecuteAsync(null);
                    }
                };
            }
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"ParameterConfigPage ViewModel init failed: {ex}");
        }
    }

    /// <summary>
    /// Attaches keyword-contains filtering to an editable ComboBox.
    /// When the user types text, the dropdown filters to items containing that text (case-insensitive).
    /// </summary>
    private void SetupFilterableComboBox(ComboBox comboBox, string displayProperty)
    {
        comboBox.AddHandler(TextBoxBase.TextChangedEvent, new TextChangedEventHandler((sender, e) =>
        {
            if (_isUpdatingFromSelection) return;

            var textBox = comboBox.Template.FindName("PART_EditableTextBox", comboBox) as TextBox;
            if (textBox == null) return;

            var filterText = textBox.Text ?? "";
            var view = CollectionViewSource.GetDefaultView(comboBox.ItemsSource);
            if (view == null) return;

            if (string.IsNullOrEmpty(filterText))
            {
                view.Filter = null;
            }
            else
            {
                view.Filter = item =>
                {
                    var value = GetPropertyValue(item, displayProperty);
                    return value != null &&
                           value.IndexOf(filterText, StringComparison.OrdinalIgnoreCase) >= 0;
                };
            }

            view.Refresh();
            comboBox.IsDropDownOpen = true;
        }));

        comboBox.SelectionChanged += (sender, e) =>
        {
            if (comboBox.SelectedItem != null)
            {
                _isUpdatingFromSelection = true;
                try
                {
                    var view = CollectionViewSource.GetDefaultView(comboBox.ItemsSource);
                    if (view != null)
                    {
                        view.Filter = null;
                        view.Refresh();
                    }
                }
                finally
                {
                    _isUpdatingFromSelection = false;
                }
            }
        };

        comboBox.DropDownClosed += (sender, e) =>
        {
            if (comboBox.SelectedItem == null) return;
            _isUpdatingFromSelection = true;
            try
            {
                var view = CollectionViewSource.GetDefaultView(comboBox.ItemsSource);
                if (view != null)
                {
                    view.Filter = null;
                    view.Refresh();
                }
            }
            finally
            {
                _isUpdatingFromSelection = false;
            }
        };
    }

    private static string? GetPropertyValue(object item, string propertyName)
    {
        return item switch
        {
            OdooProduct product => propertyName == "Name" ? product.Name : null,
            OdooSetupValue setup => propertyName == "Value" ? setup.Value : null,
            OdooColor color => propertyName == "Name" ? color.Name : null,
            _ => item.GetType().GetProperty(propertyName)?.GetValue(item)?.ToString()
        };
    }
}
