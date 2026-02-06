# US-003-12: Persist Credentials

## User Story
**As a** CAD Engineer,
**I want to** have connection credentials persisted securely between sessions,
**So that** I do not have to re-enter credentials every launch.

## Parent Feature
- **FR**: [FR-003-odoo-connection](../FR-003-odoo-connection/FR-003-odoo-connection.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: The application provides a "Remember Me" or "Save Credentials" option on the connection settings form
- [ ] AC-02: When enabled, credentials (username and password) are persisted securely using Windows DPAPI or Windows Credential Manager
- [ ] AC-03: Persisted credentials are loaded and populated into the form fields on application launch
- [ ] AC-04: The password is never stored in `appsettings.json` in plaintext
- [ ] AC-05: The Password field uses masked input (PasswordBox) regardless of whether credentials are persisted
- [ ] AC-06: The user can clear persisted credentials by unchecking the "Remember Me" option or explicitly deleting them
- [ ] AC-07: If credential retrieval fails (e.g., credential store corrupted), the application starts with empty password field and logs a warning
- [ ] AC-08: Persisted credentials are scoped to the specific Server URL and Database combination

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-003-009 | The Password/API Key field SHALL use PasswordBox (masked input) and SHALL NOT be stored in `appsettings.json` in plaintext | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-003-12-01 | Add "Remember Me" checkbox to the connection settings form in XAML | `Views/Pages/OdooConnectionPage.xaml` | S |
| TASK-003-12-02 | Create ICredentialService interface with SaveCredentials(), LoadCredentials(), ClearCredentials() methods | `OdooAutoCAD.Core/Odoo/IOdooService.cs` | S |
| TASK-003-12-03 | Implement CredentialService using Windows DPAPI (ProtectedData) for secure credential encryption | `OdooAutoCAD.Core/Odoo/OdooService.cs` | L |
| TASK-003-12-04 | Alternative: Implement CredentialService using Windows Credential Manager (CredentialManager NuGet) | `OdooAutoCAD.Core/Odoo/OdooService.cs` | L |
| TASK-003-12-05 | Load persisted credentials on ViewModel initialization and populate form fields | `ViewModels/OdooConnectionViewModel.cs` | M |
| TASK-003-12-06 | Save credentials on successful connection when "Remember Me" is checked | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-12-07 | Clear persisted credentials when "Remember Me" is unchecked | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-12-08 | Add error handling for credential retrieval failures with logging and graceful fallback | `ViewModels/OdooConnectionViewModel.cs` | S |

## Dependencies
- Depends on: US-003-01
- Blocks: None

## Notes
- The FR document Section 7 references an optional `ICredentialService` with `SaveCredentials()` and `LoadCredentials()` methods for secure credential persistence via Windows Credential Manager or DPAPI.
- Two implementation approaches are available:
  1. **Windows DPAPI** (`System.Security.Cryptography.ProtectedData`): Encrypts data tied to the current Windows user. Simpler but data is lost if the user profile is deleted.
  2. **Windows Credential Manager**: Uses the OS credential store. More robust and visible in Windows Credential Manager UI.
- Credentials should be scoped by Server URL + Database combination so different environments do not overwrite each other.
- The Python version stores a UUID token in `token.yaml`; the C# version stores username + password since it uses session-based authentication.
- Security consideration: Even with DPAPI, the password is accessible to any process running as the same Windows user. This is acceptable for desktop application credentials but should be documented.
