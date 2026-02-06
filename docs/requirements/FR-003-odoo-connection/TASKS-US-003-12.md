# TASKS: US-003-12 — Persist Credentials

> **Parent US**: [US-003-12](US-003-12-persist-credentials.md)
> **Parent FR**: [FR-003](FR-003-odoo-connection.md)
> **Priority**: P2
> **Tasks**: 8 | **Effort**: 4S + 2M + 2L
> **Status**: Not Started

## Prerequisites
- [ ] US-003-01 (credential input form) must be completed so credential fields exist

## Acceptance Criteria
- [ ] AC-01: The application provides a "Remember Me" or "Save Credentials" option on the connection settings form
- [ ] AC-02: When enabled, credentials (username and password) are persisted securely using Windows DPAPI or Windows Credential Manager
- [ ] AC-03: Persisted credentials are loaded and populated into the form fields on application launch
- [ ] AC-04: The password is never stored in `appsettings.json` in plaintext
- [ ] AC-05: The Password field uses masked input (PasswordBox) regardless of whether credentials are persisted
- [ ] AC-06: The user can clear persisted credentials by unchecking the "Remember Me" option or explicitly deleting them
- [ ] AC-07: If credential retrieval fails (e.g., credential store corrupted), the application starts with empty password field and logs a warning
- [ ] AC-08: Persisted credentials are scoped to the specific Server URL and Database combination

---

## TASK-003-12-01: Add "Remember Me" checkbox to XAML

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a `CheckBox` labeled "Remember Me" below the Password field or in the button bar area:
  ```xml
  <CheckBox Content="Remember Me"
            IsChecked="{Binding RememberCredentials, Mode=TwoWay}"
            Margin="0,5" />
  ```
- Position it near the Password `PasswordBox` so the relationship is clear
- The checkbox should be enabled regardless of connection state (user can choose to remember before connecting)
- Add a tooltip: "Save credentials securely for this server and database"

### How to verify
- [ ] "Remember Me" checkbox is visible on the connection settings form (AC-01)
- [ ] Checkbox is bound to RememberCredentials property (AC-01)

---

## TASK-003-12-02: Create ICredentialService interface

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Services/ICredentialService.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-12-03, TASK-003-12-04 |

### What to do
- Create `ICredentialService` interface in namespace `OdooAutoCAD.Core.Services`:
  ```csharp
  public interface ICredentialService
  {
      /// <summary>
      /// Saves credentials securely, scoped to the server URL and database.
      /// </summary>
      Task SaveCredentialsAsync(string serverUrl, string database, string username, string password);

      /// <summary>
      /// Loads persisted credentials for the specified server URL and database.
      /// Returns null if no credentials are stored.
      /// </summary>
      Task<StoredCredential?> LoadCredentialsAsync(string serverUrl, string database);

      /// <summary>
      /// Clears persisted credentials for the specified server URL and database.
      /// </summary>
      Task ClearCredentialsAsync(string serverUrl, string database);

      /// <summary>
      /// Checks if credentials exist for the specified server URL and database.
      /// </summary>
      Task<bool> HasCredentialsAsync(string serverUrl, string database);
  }

  public record StoredCredential(string Username, string Password, string ServerUrl, string Database);
  ```
- The credential key should be composed from `serverUrl + database` to scope per-environment
- Use a consistent key format: e.g., `"OdooAutoCAD:{serverUrl}:{database}"`

### How to verify
- [ ] ICredentialService interface declares SaveCredentials, LoadCredentials, ClearCredentials methods (AC-02)
- [ ] Credentials are scoped to Server URL and Database combination (AC-08)

---

## TASK-003-12-03: Implement CredentialService using Windows DPAPI

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Services/DpapiCredentialService.cs` |
| Estimate | L |
| Depends On | TASK-003-12-02 |
| Blocks | TASK-003-12-05 |

### What to do
- Create `DpapiCredentialService` implementing `ICredentialService` in namespace `OdooAutoCAD.Core.Services`:
  ```csharp
  using System.Security.Cryptography;
  using System.Text;
  using System.Text.Json;

  public class DpapiCredentialService : ICredentialService
  {
      private readonly string _credentialDirectory;
      private readonly ILogger<DpapiCredentialService>? _logger;

      public DpapiCredentialService(ILogger<DpapiCredentialService>? logger = null)
      {
          _credentialDirectory = Path.Combine(
              Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
              "OdooAutoCAD", "credentials");
          Directory.CreateDirectory(_credentialDirectory);
          _logger = logger;
      }

      public async Task SaveCredentialsAsync(string serverUrl, string database, string username, string password)
      {
          var credential = new StoredCredential(username, password, serverUrl, database);
          var json = JsonSerializer.Serialize(credential);
          var bytes = Encoding.UTF8.GetBytes(json);
          var encrypted = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
          var filePath = GetCredentialFilePath(serverUrl, database);
          await File.WriteAllBytesAsync(filePath, encrypted);
      }

      public async Task<StoredCredential?> LoadCredentialsAsync(string serverUrl, string database)
      {
          var filePath = GetCredentialFilePath(serverUrl, database);
          if (!File.Exists(filePath)) return null;
          var encrypted = await File.ReadAllBytesAsync(filePath);
          var decrypted = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
          var json = Encoding.UTF8.GetString(decrypted);
          return JsonSerializer.Deserialize<StoredCredential>(json);
      }

      public Task ClearCredentialsAsync(string serverUrl, string database)
      {
          var filePath = GetCredentialFilePath(serverUrl, database);
          if (File.Exists(filePath)) File.Delete(filePath);
          return Task.CompletedTask;
      }

      private string GetCredentialFilePath(string serverUrl, string database)
      {
          var key = Convert.ToBase64String(
              SHA256.HashData(Encoding.UTF8.GetBytes($"{serverUrl}:{database}")))
              .Replace("/", "_").Replace("+", "-");
          return Path.Combine(_credentialDirectory, $"{key}.dat");
      }
  }
  ```
- `DataProtectionScope.CurrentUser` ties encryption to the current Windows user profile
- Store encrypted credential files in `%LOCALAPPDATA%\OdooAutoCAD\credentials\`
- Use SHA256 hash of `serverUrl:database` as the filename to scope per environment (AC-08)
- Add NuGet reference: `System.Security.Cryptography.ProtectedData` package

### How to verify
- [ ] Credentials are encrypted using Windows DPAPI (AC-02)
- [ ] Encrypted data is tied to current Windows user (AC-02)
- [ ] Password is never stored in plaintext (AC-04)
- [ ] Credentials are scoped to server URL and database (AC-08)

---

## TASK-003-12-04: Alternative: Implement CredentialService using Windows Credential Manager

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Services/WindowsCredentialService.cs` |
| Estimate | L |
| Depends On | TASK-003-12-02 |
| Blocks | TASK-003-12-05 |

### What to do
- Create `WindowsCredentialService` implementing `ICredentialService` as an alternative to DPAPI:
  ```csharp
  using AdysTech.CredentialManager; // or CredentialManagement NuGet

  public class WindowsCredentialService : ICredentialService
  {
      private const string CredentialPrefix = "OdooAutoCAD";

      public Task SaveCredentialsAsync(string serverUrl, string database, string username, string password)
      {
          var target = $"{CredentialPrefix}:{serverUrl}:{database}";
          var credential = new NetworkCredential(username, password);
          CredentialManager.SaveCredentials(target, credential);
          return Task.CompletedTask;
      }

      public Task<StoredCredential?> LoadCredentialsAsync(string serverUrl, string database)
      {
          var target = $"{CredentialPrefix}:{serverUrl}:{database}";
          var credential = CredentialManager.GetCredentials(target);
          if (credential == null) return Task.FromResult<StoredCredential?>(null);
          return Task.FromResult<StoredCredential?>(
              new StoredCredential(credential.UserName, credential.Password, serverUrl, database));
      }

      public Task ClearCredentialsAsync(string serverUrl, string database)
      {
          var target = $"{CredentialPrefix}:{serverUrl}:{database}";
          CredentialManager.RemoveCredentials(target);
          return Task.CompletedTask;
      }
  }
  ```
- Add NuGet package: `AdysTech.CredentialManager` or `CredentialManagement`
- Credentials are visible in Windows Credential Manager UI (Control Panel > Credential Manager > Windows Credentials)
- This approach is more robust than DPAPI file-based storage and survives profile migrations
- Only one of TASK-003-12-03 or TASK-003-12-04 needs to be implemented; choose based on team preference

### How to verify
- [ ] Credentials are stored in Windows Credential Manager (AC-02)
- [ ] Credentials are visible in Credential Manager UI for management (AC-06)
- [ ] Scoped by server URL and database (AC-08)

---

## TASK-003-12-05: Load persisted credentials on ViewModel initialization

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | M |
| Depends On | TASK-003-12-03 or TASK-003-12-04 |
| Blocks | TASK-003-12-06 |

### What to do
- Inject `ICredentialService` in the ViewModel constructor
- Add observable property: `private bool _rememberCredentials;`
- After loading configuration defaults (from `appsettings.json`), attempt to load persisted credentials:
  ```csharp
  private async Task LoadPersistedCredentialsAsync()
  {
      try
      {
          if (string.IsNullOrEmpty(ServerUrl) || string.IsNullOrEmpty(Database)) return;

          var hasCredentials = await _credentialService.HasCredentialsAsync(ServerUrl, Database);
          if (hasCredentials)
          {
              var stored = await _credentialService.LoadCredentialsAsync(ServerUrl, Database);
              if (stored != null)
              {
                  Username = stored.Username;
                  Password = stored.Password;
                  RememberCredentials = true;
              }
          }
      }
      catch (Exception ex)
      {
          _logger?.LogWarning(ex, "Failed to load persisted credentials. Starting with empty password.");
          Password = string.Empty;
          RememberCredentials = false;
      }
  }
  ```
- Call `LoadPersistedCredentialsAsync()` during ViewModel initialization (after configuration defaults are set)
- Handle `CryptographicException`, `IOException`, and other failures gracefully (AC-07)
- When `ServerUrl` or `Database` changes, attempt to load credentials for the new combination

### How to verify
- [ ] Persisted credentials are loaded on launch (AC-03)
- [ ] Password field is populated from credential store (AC-03)
- [ ] Password field uses masked input regardless (AC-05)
- [ ] Credential retrieval failure results in empty password with warning log (AC-07)

---

## TASK-003-12-06: Save credentials on successful connection

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-003-12-05 |
| Blocks | TASK-003-12-07 |

### What to do
- In `ConnectAsync()`, after successful connection, check if "Remember Me" is enabled:
  ```csharp
  if (RememberCredentials)
  {
      await _credentialService.SaveCredentialsAsync(
          ServerUrl.TrimEnd('/'), Database, Username, Password);
      _logger?.LogInformation("Credentials saved for {ServerUrl}/{Database}", ServerUrl, Database);
  }
  ```
- Only save after a SUCCESSFUL connection (validates that the credentials are correct)
- Do not save if `RememberCredentials` is false
- Log the save operation (but do NOT log the password or any credential details)

### How to verify
- [ ] Credentials saved when "Remember Me" is checked and connection succeeds (AC-02)
- [ ] Credentials not saved when "Remember Me" is unchecked (AC-01)

---

## TASK-003-12-07: Clear persisted credentials when "Remember Me" is unchecked

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-003-12-06 |
| Blocks | None |

### What to do
- Implement `partial void OnRememberCredentialsChanged(bool value)`:
  ```csharp
  partial void OnRememberCredentialsChanged(bool value)
  {
      if (!value && !string.IsNullOrEmpty(ServerUrl) && !string.IsNullOrEmpty(Database))
      {
          _ = _credentialService.ClearCredentialsAsync(ServerUrl.TrimEnd('/'), Database);
          _logger?.LogInformation("Persisted credentials cleared for {ServerUrl}/{Database}", ServerUrl, Database);
      }
  }
  ```
- When the user unchecks "Remember Me", immediately clear the stored credentials
- Fire-and-forget the clear operation (use `_ =` pattern or wrap in try-catch)
- Also clear credentials if the user explicitly clicks a "Clear Credentials" button (optional)
- Consider adding a confirmation dialog before clearing: "Are you sure you want to remove saved credentials?"

### How to verify
- [ ] Unchecking "Remember Me" clears persisted credentials (AC-06)
- [ ] Credentials are removed from the secure store (AC-06)

---

## TASK-003-12-08: Add error handling for credential retrieval failures

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In `LoadPersistedCredentialsAsync()`, handle specific exception types:
  - `CryptographicException`: credential store corrupted or user profile changed
  - `IOException`: credential file inaccessible
  - `UnauthorizedAccessException`: permission denied
  - `JsonException`: stored data format is invalid
- For all failure cases:
  - Log warning with exception details: `_logger.LogWarning(ex, "Failed to retrieve stored credentials...")`
  - Set `Password = string.Empty` (do not display stale data)
  - Set `RememberCredentials = false` (uncheck the box since credentials are unavailable)
  - Do NOT show a user-facing error dialog (silent fallback per AC-07)
- In `SaveCredentialsAsync()`, also handle and log failures gracefully without blocking the connection flow

### How to verify
- [ ] Credential retrieval failures result in empty password and warning log (AC-07)
- [ ] Application starts normally even with corrupted credential store (AC-07)

---

## Dependency Graph
```
TASK-003-12-02 (ICredentialService interface)
    |
    +---> TASK-003-12-03 (DPAPI implementation)
    |         |
    +---> TASK-003-12-04 (Windows Credential Manager -- alternative)
              |
              +---> TASK-003-12-05 (Load on initialization)
                        |
                        +---> TASK-003-12-06 (Save on connect)
                                  |
                                  +---> TASK-003-12-07 (Clear on uncheck)

TASK-003-12-01 (XAML Remember Me checkbox)

TASK-003-12-08 (Error handling)
```
