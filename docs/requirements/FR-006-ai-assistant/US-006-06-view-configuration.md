# US-006-06: View Configuration

## User Story
**As a** System Admin,
**I want to** see the server port and endpoint configuration,
**So that** I can configure AI assistant clients to connect to the correct address.

## Parent Feature
- **FR**: [FR-006-ai-assistant](../FR-006-ai-assistant/FR-006-ai-assistant.md)
- **Priority**: P3

## Acceptance Criteria
- [ ] AC-01: The page displays the configured server port (default: 8084)
- [ ] AC-02: The page displays the SSE endpoint URL (e.g., http://localhost:8084/sse)
- [ ] AC-03: The page displays the Health endpoint URL (e.g., http://localhost:8084/health)
- [ ] AC-04: The page displays the JSON-RPC messages endpoint URL (e.g., http://localhost:8084/messages)
- [ ] AC-05: The page displays the transport mode (SSE)
- [ ] AC-06: The page displays configuration guidance for AI assistant clients (Gemini CLI, Claude Code)
- [ ] AC-07: Endpoint URLs automatically update when the server port is changed

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-006-045 | Page SHALL display the configured server port (default: 8084) | Must |
| FR-006-046 | Page SHALL allow the user to change the server port when the server is stopped | Should |
| FR-006-047 | Page SHALL display the SSE endpoint URL | Should |
| FR-006-048 | Page SHALL display configuration guidance for AI assistant clients | Could |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-006-06-01 | Create Configuration panel XAML layout with server settings, endpoint URLs, and client guidance sections | `Views/Pages/MCPAssistantPage.xaml` | M |
| TASK-006-06-02 | Add ViewModel properties: SSEEndpointUrl, HealthEndpointUrl, MessagesEndpointUrl computed from ServerPort | `ViewModels/MCPViewModel.cs` | S |
| TASK-006-06-03 | Display GUI Proxy status (running/stopped) and pending request count in Configuration panel | `Views/Pages/MCPAssistantPage.xaml` | S |
| TASK-006-06-04 | Display prerequisite legend ([A] = AutoCAD required, [O] = Odoo required) in Configuration panel | `Views/Pages/MCPAssistantPage.xaml` | S |
| TASK-006-06-05 | Add client configuration guidance text for Gemini CLI and Claude Code setup | `Views/Pages/MCPAssistantPage.xaml` | S |

## Dependencies
- Depends on: None (configuration is always viewable)
- Blocks: None

## Notes
- The Configuration panel is positioned in the right-center area of the MCP Assistant page.
- Endpoint URLs are computed properties derived from ServerPort: `http://localhost:{ServerPort}/sse`, `http://localhost:{ServerPort}/health`, `http://localhost:{ServerPort}/messages`.
- The panel also shows GUI Proxy status information including IsGUIProxyRunning, PendingRequestCount, and optionally average execution time from ProxyStatistics.
- The prerequisite legend explains the [A] and [O] badges used in the tool registry panel.
- Client guidance includes the Gemini CLI MCP configuration path and connection settings format.
