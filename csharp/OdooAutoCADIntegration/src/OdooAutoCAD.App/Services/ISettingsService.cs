using OdooAutoCAD.Configuration;

namespace OdooAutoCAD.App.Services;

/// <summary>
/// Abstraction for reading/writing application settings.
/// Wraps AppDbContext (ServerConfigs, UserPreferences) and ConfigurationLoader.
/// </summary>
public interface ISettingsService
{
    Task<Dictionary<string, string?>> LoadServerConfigsAsync();
    Task SaveServerConfigAsync(string key, string? value, string? description = null);
    Task SaveServerConfigsAsync(Dictionary<string, string?> configs);
    Task<string?> GetPreferenceAsync(string key, string? defaultValue = null);
    Task SetPreferenceAsync(string key, string? value, string dataType = "string");
    AppSettings GetAppSettings();
    Task SaveAppSettingsAsync(AppSettings settings);
}
