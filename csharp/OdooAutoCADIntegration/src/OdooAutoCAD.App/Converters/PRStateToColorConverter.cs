using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace OdooAutoCAD.App.Converters;

/// <summary>
/// Converts a PR state string to a color brush for state badges.
/// </summary>
public class PRStateToColorConverter : IValueConverter
{
    private static readonly SolidColorBrush GrayBrush =
        new((Color)ColorConverter.ConvertFromString("#757575"));
    private static readonly SolidColorBrush BlueBrush =
        new((Color)ColorConverter.ConvertFromString("#1A73E8"));
    private static readonly SolidColorBrush GreenBrush =
        new((Color)ColorConverter.ConvertFromString("#388E3C"));
    private static readonly SolidColorBrush RedBrush =
        new((Color)ColorConverter.ConvertFromString("#D32F2F"));

    static PRStateToColorConverter()
    {
        GrayBrush.Freeze();
        BlueBrush.Freeze();
        GreenBrush.Freeze();
        RedBrush.Freeze();
    }

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value as string switch
        {
            "submitted" or "sent" => BlueBrush,
            "approved" or "done" => GreenBrush,
            "rejected" or "cancel" => RedBrush,
            _ => GrayBrush // draft and unknown
        };
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
