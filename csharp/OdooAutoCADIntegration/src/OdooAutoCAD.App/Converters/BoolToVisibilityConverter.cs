using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace OdooAutoCAD.App.Converters;

/// <summary>
/// Converts a boolean to Visibility (true → Visible, false → Collapsed).
/// Use ConverterParameter="Invert" to reverse the logic.
/// </summary>
public class BoolToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var boolValue = value is true;
        var invert = string.Equals(parameter?.ToString(), "Invert", StringComparison.OrdinalIgnoreCase);

        if (invert)
            boolValue = !boolValue;

        return boolValue ? Visibility.Visible : Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
