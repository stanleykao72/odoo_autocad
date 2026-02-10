using System.Globalization;
using System.Windows.Data;

namespace OdooAutoCAD.App.Converters;

/// <summary>
/// Multi-value converter that returns true when the first two binding values are equal strings.
/// Used for active navigation button DataTrigger.
/// </summary>
public class EqualityConverter : IMultiValueConverter
{
    public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
    {
        if (values.Length < 2) return false;
        return string.Equals(values[0]?.ToString(), values[1]?.ToString(), StringComparison.Ordinal);
    }

    public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
    {
        throw new NotSupportedException();
    }
}
