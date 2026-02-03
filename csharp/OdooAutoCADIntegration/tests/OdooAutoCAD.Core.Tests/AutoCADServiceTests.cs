// OdooAutoCAD.Core.Tests/AutoCADServiceTests.cs
// Unit tests for AutoCAD service

using FluentAssertions;
using Moq;
using OdooAutoCAD.Core.AutoCAD;
using Xunit;

namespace OdooAutoCAD.Core.Tests;

public class AutoCADServiceTests
{
    [Fact]
    public void NewService_ShouldNotBeConnected()
    {
        // Arrange & Act
        using var service = new AutoCADService();

        // Assert
        service.IsConnected.Should().BeFalse();
    }

    [Fact]
    public void Point3D_ShouldStoreCoordinates()
    {
        // Arrange & Act
        var point = new Point3D(10.5, 20.5, 30.5);

        // Assert
        point.X.Should().Be(10.5);
        point.Y.Should().Be(20.5);
        point.Z.Should().Be(30.5);
    }

    [Fact]
    public void Point3D_WithDefaultZ_ShouldBeZero()
    {
        // Arrange & Act
        var point = new Point3D(10, 20);

        // Assert
        point.Z.Should().Be(0);
    }

    [Fact]
    public void LayoutInfo_ShouldStoreProperties()
    {
        // Arrange & Act
        var layout = new LayoutInfo(
            Name: "Layout1",
            TabOrder: 1,
            IsModelSpace: false,
            PlotConfigurationName: "DWG To PDF.pc3");

        // Assert
        layout.Name.Should().Be("Layout1");
        layout.TabOrder.Should().Be(1);
        layout.IsModelSpace.Should().BeFalse();
        layout.PlotConfigurationName.Should().Be("DWG To PDF.pc3");
    }

    [Fact]
    public void AutoCADStatus_WhenNotConnected_ShouldHaveErrorMessage()
    {
        // Arrange & Act
        var status = new AutoCADStatus(
            IsConnected: false,
            ApplicationName: null,
            Version: null,
            CurrentDocument: null,
            OpenDocuments: null,
            ErrorMessage: "Not connected to AutoCAD");

        // Assert
        status.IsConnected.Should().BeFalse();
        status.ErrorMessage.Should().NotBeNullOrEmpty();
    }

    [Fact]
    public void LayoutData_ShouldInitializeWithEmptyCollections()
    {
        // Arrange & Act
        var data = new LayoutData();

        // Assert
        data.Parameters.Should().NotBeNull().And.BeEmpty();
        data.Tables.Should().NotBeNull().And.BeEmpty();
        data.LayoutName.Should().BeEmpty();
    }

    [Fact]
    public void TableData_ShouldInitializeWithEmptyCells()
    {
        // Arrange & Act
        var table = new TableData();

        // Assert
        table.Cells.Should().NotBeNull().And.BeEmpty();
        table.RowCount.Should().Be(0);
        table.ColumnCount.Should().Be(0);
    }
}
