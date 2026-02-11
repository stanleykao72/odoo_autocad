using FluentAssertions;
using System.Globalization;
using System.Windows.Media;
using OdooAutoCAD.App.Converters;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

public class PRStateToColorConverterTests
{
    private readonly PRStateToColorConverter _converter = new();

    [Fact]
    public void Convert_Draft_ReturnsGray()
    {
        var result = _converter.Convert("draft", typeof(Brush), null!, CultureInfo.InvariantCulture);

        result.Should().BeOfType<SolidColorBrush>();
        var brush = (SolidColorBrush)result;
        brush.Color.ToString().Should().Be("#FF757575");
    }

    [Fact]
    public void Convert_Submitted_ReturnsBlue()
    {
        var result = _converter.Convert("submitted", typeof(Brush), null!, CultureInfo.InvariantCulture);

        result.Should().BeOfType<SolidColorBrush>();
        var brush = (SolidColorBrush)result;
        brush.Color.ToString().Should().Be("#FF1A73E8");
    }

    [Fact]
    public void Convert_Approved_ReturnsGreen()
    {
        var result = _converter.Convert("approved", typeof(Brush), null!, CultureInfo.InvariantCulture);

        result.Should().BeOfType<SolidColorBrush>();
        var brush = (SolidColorBrush)result;
        brush.Color.ToString().Should().Be("#FF388E3C");
    }

    [Fact]
    public void Convert_Rejected_ReturnsRed()
    {
        var result = _converter.Convert("rejected", typeof(Brush), null!, CultureInfo.InvariantCulture);

        result.Should().BeOfType<SolidColorBrush>();
        var brush = (SolidColorBrush)result;
        brush.Color.ToString().Should().Be("#FFD32F2F");
    }
}
