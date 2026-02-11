namespace OdooAutoCAD.App.Services;

/// <summary>
/// Service for securely persisting user credentials using DPAPI.
/// </summary>
public interface ICredentialService
{
    Task SaveCredentialsAsync(string swaggerUrl, string userToken);
    Task<(string? SwaggerUrl, string? UserToken)> LoadCredentialsAsync();
    Task ClearCredentialsAsync();
    bool HasSavedCredentials { get; }
}
