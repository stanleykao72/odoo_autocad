using OdooAutoCAD.Core.AutoCAD;
using Xunit;

namespace OdooAutoCAD.Integration.Tests.Services;

public class DrawingDataServiceEnumTests
{
    [Fact]
    public void AutoCADOperationMode_HasExpectedValues()
    {
        Assert.Equal(0, (int)AutoCADOperationMode.COM);
        Assert.Equal(1, (int)AutoCADOperationMode.File);
    }

    [Fact]
    public void WriteStrategy_HasExpectedValues()
    {
        Assert.Equal(0, (int)WriteStrategy.DirectDwg);
        Assert.Equal(1, (int)WriteStrategy.ExportDxf);
        Assert.Equal(2, (int)WriteStrategy.SidecarJson);
        Assert.Equal(3, (int)WriteStrategy.WriteUnsupported);
    }

    [Fact]
    public void WritebackResult_RecordEquality()
    {
        var a = new WritebackResult(true, "OK", WriteStrategy.DirectDwg);
        var b = new WritebackResult(true, "OK", WriteStrategy.DirectDwg);
        Assert.Equal(a, b);
    }

    [Fact]
    public void WritebackResult_WithOutputPath()
    {
        var result = new WritebackResult(true, "Saved", WriteStrategy.SidecarJson, "/tmp/file.json");
        Assert.True(result.Success);
        Assert.Equal("Saved", result.Message);
        Assert.Equal(WriteStrategy.SidecarJson, result.StrategyUsed);
        Assert.Equal("/tmp/file.json", result.OutputPath);
    }

    [Fact]
    public void WritebackResult_DefaultOutputPath_IsNull()
    {
        var result = new WritebackResult(false, "Error", WriteStrategy.WriteUnsupported);
        Assert.Null(result.OutputPath);
    }
}
