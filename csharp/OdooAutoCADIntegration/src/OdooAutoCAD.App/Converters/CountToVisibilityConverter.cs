using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace OdooAutoCAD.App.Converters;

/// <summary>
/// Converts a count (int) to Visibility.
/// Returns Visible if count > 0, otherwise Collapsed.
/// </summary>
public class CountToVisibilityConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
    {
        if (value is int count)
        {
            return count > 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        return Visibility.Collapsed;
    }

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
    {
        throw new NotImplementedException();
    }
}
