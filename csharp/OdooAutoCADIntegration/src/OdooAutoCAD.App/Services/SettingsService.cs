using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using OdooAutoCAD.Configuration;
using OdooAutoCAD.Data.Context;
using OdooAutoCAD.Data.Entities;

namespace OdooAutoCAD.App.Services;

/// <summary>
/// Settings service implementation.
/// Reads/writes ServerConfigs and UserPreferences via EF Core,
/// and delegates JSON/YAML persistence to ConfigurationLoader.
/// </summary>
public class SettingsService : ISettingsService
{
    private readonly AppDbContextFactory _dbFactory;
    private readonly ConfigurationLoader _configLoader;
    private readonly ILogger<SettingsService>? _logger;

    public SettingsService(
        AppDbContextFactory dbFactory,
        ConfigurationLoader configLoader,
        ILogger<SettingsService>? logger = null)
    {
        _dbFactory = dbFactory;
        _configLoader = configLoader;
        _logger = logger;
    }

    public async Task<Dictionary<string, string?>> LoadServerConfigsAsync()
    {
        await using var db = _dbFactory.CreateDbContext();
        await db.Database.EnsureCreatedAsync();

        var configs = await db.ServerConfigs.ToListAsync();
        return configs.ToDictionary(c => c.ConfigKey, c => c.ConfigValue);
    }

    public async Task SaveServerConfigAsync(string key, string? value, string? description = null)
    {
        await using var db = _dbFactory.CreateDbContext();
        await db.Database.EnsureCreatedAsync();

        var existing = await db.ServerConfigs.FirstOrDefaultAsync(c => c.ConfigKey == key);
        if (existing != null)
        {
            existing.ConfigValue = value;
            existing.UpdatedAt = DateTime.UtcNow;
            if (description != null)
                existing.Description = description;
        }
        else
        {
            db.ServerConfigs.Add(new ServerConfig
            {
                ConfigKey = key,
                ConfigValue = value,
                Description = description,
                CreatedAt = DateTime.UtcNow,
                UpdatedAt = DateTime.UtcNow
            });
        }

        await db.SaveChangesAsync();
        _logger?.LogDebug("Saved server config: {Key}", key);
    }

    public async Task SaveServerConfigsAsync(Dictionary<string, string?> configs)
    {
        await using var db = _dbFactory.CreateDbContext();
        await db.Database.EnsureCreatedAsync();

        foreach (var (key, value) in configs)
        {
            var existing = await db.ServerConfigs.FirstOrDefaultAsync(c => c.ConfigKey == key);
            if (existing != null)
            {
                existing.ConfigValue = value;
                existing.UpdatedAt = DateTime.UtcNow;
            }
            else
            {
                db.ServerConfigs.Add(new ServerConfig
                {
                    ConfigKey = key,
                    ConfigValue = value,
                    CreatedAt = DateTime.UtcNow,
                    UpdatedAt = DateTime.UtcNow
                });
            }
        }

        await db.SaveChangesAsync();
        _logger?.LogDebug("Saved {Count} server configs", configs.Count);
    }

    public async Task<string?> GetPreferenceAsync(string key, string? defaultValue = null)
    {
        await using var db = _dbFactory.CreateDbContext();
        await db.Database.EnsureCreatedAsync();

        var pref = await db.UserPreferences.FirstOrDefaultAsync(p => p.PreferenceKey == key);
        return pref?.PreferenceValue ?? defaultValue;
    }

    public async Task SetPreferenceAsync(string key, string? value, string dataType = "string")
    {
        await using var db = _dbFactory.CreateDbContext();
        await db.Database.EnsureCreatedAsync();

        var existing = await db.UserPreferences.FirstOrDefaultAsync(p => p.PreferenceKey == key);
        if (existing != null)
        {
            existing.PreferenceValue = value;
            existing.DataType = dataType;
            existing.UpdatedAt = DateTime.UtcNow;
        }
        else
        {
            db.UserPreferences.Add(new UserPreference
            {
                PreferenceKey = key,
                PreferenceValue = value,
                DataType = dataType,
                UpdatedAt = DateTime.UtcNow
            });
        }

        await db.SaveChangesAsync();
    }

    public AppSettings GetAppSettings()
    {
        return _configLoader.GetSettings();
    }

    public async Task SaveAppSettingsAsync(AppSettings settings)
    {
        // Save non-sensitive settings to JSON
        _configLoader.SaveToJson(settings);

        // Save sensitive settings (token) to DB only
        // Token is handled separately via SaveServerConfigAsync
        await Task.CompletedTask;
    }
}
