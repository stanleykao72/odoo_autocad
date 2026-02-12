using FluentAssertions;
using OdooAutoCAD.App.Converters;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

public class DateTimeToStringConverterTests
{
    private readonly DateTimeToStringConverter _converter = new();

    [Fact]
    public void Convert_Null_ReturnsNeverSynced()
    {
        var result = _converter.Convert(null!, typeof(string), null!, System.Globalization.CultureInfo.InvariantCulture);

        result.Should().Be("Never synced");
    }

    [Fact]
    public void Convert_DateTime_ReturnsFormatted()
    {
        var dt = new DateTime(2026, 2, 12, 14, 30, 45);

        var result = _converter.Convert(dt, typeof(string), null!, System.Globalization.CultureInfo.InvariantCulture);

        result.Should().Be("2026-02-12 14:30:45");
    }
}
