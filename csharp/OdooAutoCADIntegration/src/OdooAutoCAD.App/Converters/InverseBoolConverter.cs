using System.Globalization;
using System.Windows.Data;

namespace OdooAutoCAD.App.Converters;

/// <summary>
/// Inverts a boolean value (true → false, false → true).
/// Useful for disabling controls when a checkbox is checked.
/// </summary>
public class InverseBoolConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is not true;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is not true;
    }
}
