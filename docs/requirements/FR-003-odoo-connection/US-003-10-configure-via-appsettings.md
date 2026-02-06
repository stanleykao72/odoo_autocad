# US-003-10: Configure via appsettings.json

## User Story
**As a** System Admin,
**I want to** configure Odoo connection settings in appsettings.json,
**So that** the application can start with pre-configured defaults.

## Parent Feature
- **FR**: [FR-003-odoo-connection](../FR-003-odoo-connection/FR-003-odoo-connection.md)
- **Priority**: P2

## Acceptance Criteria
- [ ] AC-01: The `appsettings.json` file supports an `Odoo` section with keys: `ServerUrl`, `Database`, `Username`, `TimeoutSeconds`
- [ ] AC-02: On application launch, the credential form fields are pre-populated with values from `appsettings.json`
- [ ] AC-03: The `TimeoutSeconds` value configures the `HttpClient.Timeout` (default: 30 seconds if not specified)
- [ ] AC-04: `TimeoutSeconds` is validated as a positive integer between 5 and 120 seconds
- [ ] AC-05: If `appsettings.json` is missing the `Odoo` section, the application starts with empty fields and default timeout
- [ ] AC-06: The Password field is intentionally excluded from `appsettings.json` for security
- [ ] AC-07: Users can override pre-configured values by typing in the UI fields at runtime
- [ ] AC-08: Configuration changes in the UI are not written back to `appsettings.json` (read-only configuration)

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-003-008 | Connection settings SHALL be pre-populated from `appsettings.json` on application launch (`Odoo:ServerUrl`, `Odoo:Database`, `Odoo:Username`) | Must |
| FR-003-010 | Connection timeout SHALL be configurable via `Odoo:TimeoutSeconds` in `appsettings.json` (default: 30 seconds) | Should |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-003-10-01 | Add Odoo configuration section to appsettings.json with ServerUrl, Database, Username, TimeoutSeconds | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` | S |
| TASK-003-10-02 | Create OdooSettings POCO class for strongly-typed configuration binding | `OdooAutoCAD.Configuration/ConfigurationLoader.cs` | S |
| TASK-003-10-03 | Inject IConfiguration and bind Odoo section in ViewModel constructor | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-10-04 | Pre-populate ServerUrl, Database, Username from configuration on ViewModel initialization | `ViewModels/OdooConnectionViewModel.cs` | S |
| TASK-003-10-05 | Configure HttpClient.Timeout from TimeoutSeconds value with validation (5-120 range, default 30) | `OdooAutoCAD.Core/Odoo/OdooService.cs` | S |
| TASK-003-10-06 | Handle missing Odoo section gracefully: start with empty fields and default timeout of 30 seconds | `ViewModels/OdooConnectionViewModel.cs` | S |

## Dependencies
- Depends on: None
- Blocks: US-003-01

## Notes
- The configuration structure in `appsettings.json` follows the pattern documented in the FR:
  ```json
  {
      "Odoo": {
          "ServerUrl": "https://odoo.example.com",
          "Database": "production",
          "Username": "",
          "TimeoutSeconds": 30
      }
  }
  ```
- This replaces the Python YAML-based configuration (`server.yaml`, `token.yaml`) with the standard .NET `IConfiguration` pattern.
- Password is intentionally excluded; see US-003-12 for secure credential persistence.
- Validation rule VR-003-007 applies: TimeoutSeconds must be between 5 and 120 seconds.
- The configuration is read-only from the application's perspective; runtime changes in the UI do not modify `appsettings.json`.
