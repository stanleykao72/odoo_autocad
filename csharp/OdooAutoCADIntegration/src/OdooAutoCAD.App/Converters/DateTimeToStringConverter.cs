using System.Globalization;
using System.Windows.Data;

namespace OdooAutoCAD.App.Converters;

/// <summary>
/// Converts a DateTime? to a formatted string.
/// null → "Never synced", value → "yyyy-MM-dd HH:mm:ss"
/// </summary>
public class DateTimeToStringConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is DateTime dateTime)
            return dateTime.ToString("yyyy-MM-dd HH:mm:ss");

        return "Never synced";
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
