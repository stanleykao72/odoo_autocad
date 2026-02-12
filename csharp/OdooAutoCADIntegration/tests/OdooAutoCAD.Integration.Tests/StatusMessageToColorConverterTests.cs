using FluentAssertions;
using System.Globalization;
using System.Windows.Media;
using OdooAutoCAD.App.Converters;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

public class StatusMessageToColorConverterTests
{
    private readonly StatusMessageToColorConverter _converter = new();

    [Fact]
    public void Convert_Success_ReturnsGreen()
    {
        var result = _converter.Convert("success", typeof(Brush), null!, CultureInfo.InvariantCulture);

        result.Should().BeOfType<SolidColorBrush>();
        var brush = (SolidColorBrush)result;
        brush.Color.ToString().Should().Be("#FF388E3C");
    }

    [Fact]
    public void Convert_Error_ReturnsRed()
    {
        var result = _converter.Convert("error", typeof(Brush), null!, CultureInfo.InvariantCulture);

        result.Should().BeOfType<SolidColorBrush>();
        var brush = (SolidColorBrush)result;
        brush.Color.ToString().Should().Be("#FFC62828");
    }

    [Fact]
    public void Convert_Info_ReturnsBlue()
    {
        var result = _converter.Convert("info", typeof(Brush), null!, CultureInfo.InvariantCulture);

        result.Should().BeOfType<SolidColorBrush>();
        var brush = (SolidColorBrush)result;
        brush.Color.ToString().Should().Be("#FF1A73E8");
    }

    [Fact]
    public void Convert_Unknown_ReturnsGray()
    {
        var result = _converter.Convert("something", typeof(Brush), null!, CultureInfo.InvariantCulture);

        result.Should().BeOfType<SolidColorBrush>();
        var brush = (SolidColorBrush)result;
        brush.Color.ToString().Should().Be("#FF616161");
    }

    [Fact]
    public void Convert_Null_ReturnsGray()
    {
        var result = _converter.Convert(null!, typeof(Brush), null!, CultureInfo.InvariantCulture);

        result.Should().BeOfType<SolidColorBrush>();
        var brush = (SolidColorBrush)result;
        brush.Color.ToString().Should().Be("#FF616161");
    }
}
