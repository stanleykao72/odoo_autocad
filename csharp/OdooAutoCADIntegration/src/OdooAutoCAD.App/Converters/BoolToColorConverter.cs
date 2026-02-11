using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace OdooAutoCAD.App.Converters;

/// <summary>
/// Converts a boolean to a color brush.
/// true → green (#388E3C), false → gray (#9E9E9E).
/// </summary>
public class BoolToColorConverter : IValueConverter
{
    private static readonly SolidColorBrush GreenBrush =
        new((Color)ColorConverter.ConvertFromString("#388E3C"));

    private static readonly SolidColorBrush GrayBrush =
        new((Color)ColorConverter.ConvertFromString("#9E9E9E"));

    static BoolToColorConverter()
    {
        GreenBrush.Freeze();
        GrayBrush.Freeze();
    }

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value is true ? GreenBrush : GrayBrush;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
