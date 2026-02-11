using System.Globalization;
using System.Windows.Data;

namespace OdooAutoCAD.App.Converters;

/// <summary>
/// Converts a PR state string to an icon character for state badges.
/// </summary>
public class PRStateToIconConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value as string switch
        {
            "submitted" or "sent" => "\u2709",   // ✉
            "approved" or "done" => "\u2714",     // ✔
            "rejected" or "cancel" => "\u2718",   // ✘
            _ => "\u25CB"                          // ○ (draft and unknown)
        };
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
