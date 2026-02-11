using FluentAssertions;
using Moq;
using OdooAutoCAD.App.Services;
using Xunit;

namespace OdooAutoCAD.Integration.Tests;

public class CredentialServiceTests
{
    private readonly Mock<ISettingsService> _mockSettings;
    private readonly Dictionary<string, string?> _store;

    public CredentialServiceTests()
    {
        _mockSettings = new Mock<ISettingsService>();
        _store = new Dictionary<string, string?>();

        // Simulate an in-memory settings store
        _mockSettings.Setup(s => s.LoadServerConfigsAsync())
            .ReturnsAsync(() => new Dictionary<string, string?>(_store));

        _mockSettings.Setup(s => s.SaveServerConfigsAsync(It.IsAny<Dictionary<string, string?>>()))
            .Callback<Dictionary<string, string?>>(configs =>
            {
                foreach (var kvp in configs)
                {
                    if (kvp.Value == null)
                        _store.Remove(kvp.Key);
                    else
                        _store[kvp.Key] = kvp.Value;
                }
            })
            .Returns(Task.CompletedTask);
    }

    private DpapiCredentialService CreateSUT()
    {
        return new DpapiCredentialService(_mockSettings.Object);
    }

    [Fact]
    public async Task SaveAndLoad_RoundTrips()
    {
        var sut = CreateSUT();

        await sut.SaveCredentialsAsync(
            "https://odoo.example.com/api/v1/swagger.json?token=abc&db=testdb",
            "my-secret-token");

        var (url, token) = await sut.LoadCredentialsAsync();

        url.Should().Be("https://odoo.example.com/api/v1/swagger.json?token=abc&db=testdb");
        token.Should().Be("my-secret-token");
    }

    [Fact]
    public async Task Clear_RemovesCredentials()
    {
        var sut = CreateSUT();

        await sut.SaveCredentialsAsync("https://url.com", "token123");
        sut.HasSavedCredentials.Should().BeTrue();

        await sut.ClearCredentialsAsync();
        sut.HasSavedCredentials.Should().BeFalse();

        var (url, token) = await sut.LoadCredentialsAsync();
        url.Should().BeNull();
        token.Should().BeNull();
    }

    [Fact]
    public async Task HasSavedCredentials_AfterSave_ReturnsTrue()
    {
        var sut = CreateSUT();
        sut.HasSavedCredentials.Should().BeFalse();

        await sut.SaveCredentialsAsync("https://url.com", "token");

        sut.HasSavedCredentials.Should().BeTrue();
    }

    [Fact]
    public async Task LoadCredentials_WhenEmpty_ReturnsNulls()
    {
        var sut = CreateSUT();

        var (url, token) = await sut.LoadCredentialsAsync();

        url.Should().BeNull();
        token.Should().BeNull();
    }
}
