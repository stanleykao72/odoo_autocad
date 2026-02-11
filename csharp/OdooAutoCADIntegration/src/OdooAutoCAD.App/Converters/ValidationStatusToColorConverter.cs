using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace OdooAutoCAD.App.Converters;

/// <summary>
/// Converts a validation status string to a color brush.
/// </summary>
public class ValidationStatusToColorConverter : IValueConverter
{
    private static readonly SolidColorBrush SuccessBrush =
        new((Color)ColorConverter.ConvertFromString("#388E3C"));
    private static readonly SolidColorBrush ErrorBrush =
        new((Color)ColorConverter.ConvertFromString("#C62828"));
    private static readonly SolidColorBrush WarningBrush =
        new((Color)ColorConverter.ConvertFromString("#F57F17"));
    private static readonly SolidColorBrush GrayBrush =
        new((Color)ColorConverter.ConvertFromString("#9E9E9E"));

    static ValidationStatusToColorConverter()
    {
        SuccessBrush.Freeze();
        ErrorBrush.Freeze();
        WarningBrush.Freeze();
        GrayBrush.Freeze();
    }

    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        return value as string switch
        {
            "Valid" => SuccessBrush,
            "Invalid" => ErrorBrush,
            "Skipped" => WarningBrush,
            _ => GrayBrush
        };
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
