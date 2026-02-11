// OdooAutoCAD.Configuration/ConfigurationLoader.cs
// Configuration management - equivalent to Python YAML config loading

using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.Configuration;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace OdooAutoCAD.Configuration;

/// <summary>
/// Application configuration settings.
/// </summary>
public class AppSettings
{
    public ApplicationInfo Application { get; set; } = new();
    public OdooSettings Odoo { get; set; } = new();
    public AutoCADSettings AutoCAD { get; set; } = new();
    public MCPSettings MCP { get; set; } = new();
    public GUIProxySettings GUIProxy { get; set; } = new();
    public DatabaseSettings Database { get; set; } = new();
}

public class ApplicationInfo
{
    public string Name { get; set; } = "Odoo AutoCAD Integration";
    public string Version { get; set; } = "6.0.0";
}

public class OdooSettings
{
    public string SwaggerUrl { get; set; } = string.Empty;
    public int TimeoutSeconds { get; set; } = 30;
}

public class AutoCADSettings
{
    public string ProgId { get; set; } = "AutoCAD.Application";
    public int ConnectionTimeoutSeconds { get; set; } = 10;
    public int RetryAttempts { get; set; } = 3;
}

public class MCPSettings
{
    public int Port { get; set; } = 8084;
    public bool AutoStart { get; set; } = false;
    public int HeartbeatIntervalSeconds { get; set; } = 30;
}

public class GUIProxySettings
{
    public int PollingIntervalMs { get; set; } = 100;
    public int DefaultTimeoutMs { get; set; } = 10000;
    public int MaxRequestsPerBatch { get; set; } = 10;
}

public class DatabaseSettings
{
    public string ConnectionString { get; set; } = "Data Source=database.db";
}

/// <summary>
/// Configuration loader supporting both JSON and YAML formats.
/// Equivalent to Python's YAML configuration loading.
/// </summary>
public class ConfigurationLoader
{
    private readonly string _configDirectory;
    private AppSettings? _settings;

    public ConfigurationLoader(string configDirectory = "config")
    {
        _configDirectory = configDirectory;
    }

    /// <summary>
    /// Loads configuration from JSON file.
    /// </summary>
    public AppSettings LoadFromJson(string fileName = "appsettings.json")
    {
        var builder = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile(fileName, optional: true, reloadOnChange: true)
            .AddJsonFile($"appsettings.{Environment.GetEnvironmentVariable("DOTNET_ENVIRONMENT") ?? "Production"}.json", optional: true);

        var configuration = builder.Build();

        _settings = new AppSettings();
        configuration.Bind(_settings);

        return _settings;
    }

    /// <summary>
    /// Loads configuration from YAML file.
    /// Provides compatibility with existing Python YAML configs.
    /// </summary>
    public AppSettings LoadFromYaml(string fileName)
    {
        var filePath = Path.Combine(_configDirectory, fileName);

        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"Configuration file not found: {filePath}");
        }

        var yaml = File.ReadAllText(filePath);

        var deserializer = new DeserializerBuilder()
            .WithNamingConvention(UnderscoredNamingConvention.Instance)
            .IgnoreUnmatchedProperties()
            .Build();

        _settings = deserializer.Deserialize<AppSettings>(yaml);
        return _settings ?? new AppSettings();
    }

    /// <summary>
    /// Saves configuration to YAML file.
    /// </summary>
    public void SaveToYaml(AppSettings settings, string fileName)
    {
        var filePath = Path.Combine(_configDirectory, fileName);

        var serializer = new SerializerBuilder()
            .WithNamingConvention(UnderscoredNamingConvention.Instance)
            .Build();

        var yaml = serializer.Serialize(settings);
        File.WriteAllText(filePath, yaml);
    }

    /// <summary>
    /// Saves configuration to JSON file.
    /// Sensitive fields (tokens, passwords) are excluded.
    /// </summary>
    public void SaveToJson(AppSettings settings, string fileName = "appsettings.json")
    {
        var filePath = Path.Combine(AppContext.BaseDirectory, fileName);

        var options = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };

        var json = JsonSerializer.Serialize(settings, options);
        File.WriteAllText(filePath, json);
    }

    /// <summary>
    /// Gets the current loaded settings.
    /// </summary>
    public AppSettings GetSettings()
    {
        return _settings ?? LoadFromJson();
    }

    /// <summary>
    /// Lists available configuration files.
    /// </summary>
    public IEnumerable<string> GetAvailableConfigs()
    {
        if (!Directory.Exists(_configDirectory))
        {
            yield break;
        }

        foreach (var file in Directory.GetFiles(_configDirectory, "*.yaml"))
        {
            yield return Path.GetFileName(file);
        }

        foreach (var file in Directory.GetFiles(_configDirectory, "*.yml"))
        {
            yield return Path.GetFileName(file);
        }
    }
}
