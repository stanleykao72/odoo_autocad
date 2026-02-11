using System.Security.Cryptography;
using System.Text;
using Microsoft.Extensions.Logging;

namespace OdooAutoCAD.App.Services;

/// <summary>
/// Credential service using Windows DPAPI (DataProtectionScope.CurrentUser).
/// Stores encrypted Base64 in the ServerConfig database table.
/// </summary>
public class DpapiCredentialService : ICredentialService
{
    private const string KeySwaggerUrl = "credential_swagger_url_encrypted";
    private const string KeyUserToken = "credential_user_token_encrypted";
    private static readonly byte[] Entropy = "OdooAutoCAD.Credentials.v1"u8.ToArray();

    private readonly ISettingsService _settingsService;
    private readonly ILogger<DpapiCredentialService>? _logger;
    private bool _hasSavedCredentials;

    public DpapiCredentialService(
        ISettingsService settingsService,
        ILogger<DpapiCredentialService>? logger = null)
    {
        _settingsService = settingsService;
        _logger = logger;

        // Check on construction (sync, best-effort)
        _ = CheckHasSavedCredentialsAsync();
    }

    public bool HasSavedCredentials => _hasSavedCredentials;

    public async Task SaveCredentialsAsync(string swaggerUrl, string userToken)
    {
        var encryptedUrl = Encrypt(swaggerUrl);
        var encryptedToken = Encrypt(userToken);

        await _settingsService.SaveServerConfigsAsync(new Dictionary<string, string?>
        {
            [KeySwaggerUrl] = encryptedUrl,
            [KeyUserToken] = encryptedToken
        });

        _hasSavedCredentials = true;
        _logger?.LogInformation("Credentials saved (DPAPI-encrypted)");
    }

    public async Task<(string? SwaggerUrl, string? UserToken)> LoadCredentialsAsync()
    {
        try
        {
            var configs = await _settingsService.LoadServerConfigsAsync();

            configs.TryGetValue(KeySwaggerUrl, out var encryptedUrl);
            configs.TryGetValue(KeyUserToken, out var encryptedToken);

            var url = Decrypt(encryptedUrl);
            var token = Decrypt(encryptedToken);

            _hasSavedCredentials = url != null || token != null;
            return (url, token);
        }
        catch (Exception ex)
        {
            _logger?.LogWarning(ex, "Failed to load credentials — may be from different user");
            _hasSavedCredentials = false;
            return (null, null);
        }
    }

    public async Task ClearCredentialsAsync()
    {
        await _settingsService.SaveServerConfigsAsync(new Dictionary<string, string?>
        {
            [KeySwaggerUrl] = null,
            [KeyUserToken] = null
        });

        _hasSavedCredentials = false;
        _logger?.LogInformation("Credentials cleared");
    }

    private async Task CheckHasSavedCredentialsAsync()
    {
        try
        {
            var configs = await _settingsService.LoadServerConfigsAsync();
            _hasSavedCredentials = configs.ContainsKey(KeySwaggerUrl) &&
                                   !string.IsNullOrEmpty(configs[KeySwaggerUrl]);
        }
        catch
        {
            _hasSavedCredentials = false;
        }
    }

    private static string? Encrypt(string plainText)
    {
        if (string.IsNullOrEmpty(plainText)) return null;

        var plainBytes = Encoding.UTF8.GetBytes(plainText);
        var encryptedBytes = ProtectedData.Protect(plainBytes, Entropy, DataProtectionScope.CurrentUser);
        return Convert.ToBase64String(encryptedBytes);
    }

    private static string? Decrypt(string? encryptedBase64)
    {
        if (string.IsNullOrEmpty(encryptedBase64)) return null;

        try
        {
            var encryptedBytes = Convert.FromBase64String(encryptedBase64);
            var plainBytes = ProtectedData.Unprotect(encryptedBytes, Entropy, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(plainBytes);
        }
        catch (CryptographicException)
        {
            // Different user or corrupted data
            return null;
        }
        catch (FormatException)
        {
            // Not valid Base64
            return null;
        }
    }
}
