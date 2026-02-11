using System.Globalization;
using System.Windows.Data;

namespace OdooAutoCAD.App.Converters;

/// <summary>
/// Converts a validation status string to a display icon character.
/// </summary>
public class ValidationStatusToIconConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value as string switch
        {
            "Valid" => "\u2714",    // ✔
            "Invalid" => "\u2718",  // ✘
            "Skipped" => "\u2013",  // –
            "Pending" => "\u25CB",  // ○
            _ => "\u25CB"
        };
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
