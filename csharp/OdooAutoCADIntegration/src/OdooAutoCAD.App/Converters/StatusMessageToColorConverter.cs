using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace OdooAutoCAD.App.Converters;

/// <summary>
/// Converts a status message type string to a SolidColorBrush for styled feedback.
/// "success" → Green, "error" → Red, "info" → Blue, default → Gray.
/// </summary>
public class StatusMessageToColorConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        var type = value as string ?? string.Empty;
        return type.ToLowerInvariant() switch
        {
            "success" => new SolidColorBrush(Color.FromRgb(0x38, 0x8E, 0x3C)), // Green
            "error" => new SolidColorBrush(Color.FromRgb(0xC6, 0x28, 0x28)),   // Red
            "info" => new SolidColorBrush(Color.FromRgb(0x1A, 0x73, 0xE8)),    // Blue
            _ => new SolidColorBrush(Color.FromRgb(0x61, 0x61, 0x61))          // Gray
        };
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
