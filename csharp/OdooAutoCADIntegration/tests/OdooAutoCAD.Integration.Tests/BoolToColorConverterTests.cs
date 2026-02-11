using System.Globalization;
using System.Windows.Media;
using FluentAssertions;
using OdooAutoCAD.App.Converters;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

public class BoolToColorConverterTests
{
    private readonly BoolToColorConverter _sut = new();

    [Fact]
    public void Convert_True_ReturnsGreenBrush()
    {
        var result = _sut.Convert(true, typeof(Brush), null!, CultureInfo.InvariantCulture);

        result.Should().BeOfType<SolidColorBrush>();
        var brush = (SolidColorBrush)result;
        brush.Color.ToString().Should().Be("#FF388E3C");
    }

    [Fact]
    public void Convert_False_ReturnsGrayBrush()
    {
        var result = _sut.Convert(false, typeof(Brush), null!, CultureInfo.InvariantCulture);

        result.Should().BeOfType<SolidColorBrush>();
        var brush = (SolidColorBrush)result;
        brush.Color.ToString().Should().Be("#FF9E9E9E");
    }

    [Fact]
    public void Convert_NonBool_ReturnsGrayBrush()
    {
        var result = _sut.Convert("not a bool", typeof(Brush), null!, CultureInfo.InvariantCulture);

        result.Should().BeOfType<SolidColorBrush>();
        var brush = (SolidColorBrush)result;
        brush.Color.ToString().Should().Be("#FF9E9E9E");
    }
}
