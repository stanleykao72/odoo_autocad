# TASKS: US-006-06 — View Configuration

> **Parent US**: [US-006-06](US-006-06-view-configuration.md)
> **Parent FR**: [FR-006](FR-006-ai-assistant.md)
> **Priority**: P3
> **Tasks**: 5 | **Effort**: 5S + 0M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] FR-008 (UI Framework - WPF/MVVM infrastructure must be in place)
- [ ] `MCPViewModel` created with `ServerPort` property (from US-006-01)
- [ ] `IGUIProxy` interface exists for proxy status display

## Acceptance Criteria
- [ ] AC-01: The page displays the configured server port (default: 8084)
- [ ] AC-02: The page displays the SSE endpoint URL (e.g., http://localhost:8084/sse)
- [ ] AC-03: The page displays the Health endpoint URL (e.g., http://localhost:8084/health)
- [ ] AC-04: The page displays the JSON-RPC messages endpoint URL (e.g., http://localhost:8084/messages)
- [ ] AC-05: The page displays the transport mode (SSE)
- [ ] AC-06: The page displays configuration guidance for AI assistant clients (Gemini CLI, Claude Code)
- [ ] AC-07: Endpoint URLs automatically update when the server port is changed

---

## TASK-006-06-01: Create Configuration panel XAML layout with server settings, endpoint URLs, and client guidance sections

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a "Configuration" section in the right-center area of `MCPAssistantPage.xaml` using a `Border` with header TextBlock
- Create three subsections within: "Server" settings, "Endpoints" listing, and "GUI Proxy" status
- In the "Server" subsection: display "Port:" with value bound to `{Binding ServerPort}`, "Host:" with static "localhost", "Transport:" with static "SSE"
- In the "Endpoints" subsection: display "SSE:" with value bound to `{Binding SSEEndpointUrl}`, "RPC:" with value bound to `{Binding MessagesEndpointUrl}`, "Health:" with value bound to `{Binding HealthEndpointUrl}`
- Use `TextBlock` elements with `TextWrapping="Wrap"` for URLs and monospace font for endpoint paths
- Style labels in secondary color and values in primary color for visual distinction

### How to verify
- [ ] Port number is displayed (AC-01)
- [ ] SSE endpoint URL is displayed (AC-02)
- [ ] Health endpoint URL is displayed (AC-03)
- [ ] Messages endpoint URL is displayed (AC-04)
- [ ] Transport mode "SSE" is displayed (AC-05)

---

## TASK-006-06-02: Add ViewModel properties for computed endpoint URLs

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add `string SSEEndpointUrl` computed property: `$"http://localhost:{ServerPort}/sse"`
- Add `string HealthEndpointUrl` computed property: `$"http://localhost:{ServerPort}/health"`
- Add `string MessagesEndpointUrl` computed property: `$"http://localhost:{ServerPort}/messages"`
- Add `string RootEndpointUrl` computed property: `$"http://localhost:{ServerPort}/"`
- Override or use `partial void OnServerPortChanged(int value)` (CommunityToolkit pattern) to call `OnPropertyChanged` for all computed endpoint URL properties when `ServerPort` changes
- This ensures XAML bindings update automatically when the port is modified (AC-07)

### How to verify
- [ ] SSEEndpointUrl returns correct URL with current port (AC-02)
- [ ] Endpoint URLs update when ServerPort changes (AC-07)
- [ ] All four endpoint URLs are computed correctly (AC-02, AC-03, AC-04)

---

## TASK-006-06-03: Display GUI Proxy status and pending request count in Configuration panel

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the "GUI Proxy" subsection of the Configuration panel, add: "Status:" label with value bound to `{Binding IsGUIProxyRunning}` using a `BoolToStatusConverter` (returns "Running"/"Stopped")
- Add "Pending:" label with value bound to `{Binding GUIProxyPendingRequests}`
- Add "Avg Time:" label with value bound to `{Binding GUIProxyAverageTime}` (formatted as milliseconds, e.g., "45ms")
- Add `[ObservableProperty]` fields in MCPViewModel: `bool _isGUIProxyRunning`, `int _guiProxyPendingRequests`, `string _guiProxyAverageTime` (default "--")
- Update these properties from `IGUIProxy.GetStatistics()` periodically (can share the uptime timer or proxy timer)
- Apply green/red color to the proxy status text matching the server status pattern

### How to verify
- [ ] GUI Proxy status is displayed (AC-01)
- [ ] Pending request count updates (AC-01)
- [ ] Average execution time is shown (AC-01)

---

## TASK-006-06-04: Display prerequisite legend in Configuration panel

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a "Prerequisites" subsection at the bottom of the Configuration panel
- Display three legend entries: "[A] = AutoCAD required", "[O] = Odoo required", "[A][O] = Both required"
- Use the same badge styling (colored borders) as used in the tool registry panel for visual consistency
- Add a note: "Tools with unmet prerequisites will return an error when invoked by AI assistants"
- Keep the legend compact with small font and secondary text color

### How to verify
- [ ] Prerequisite legend explains [A] and [O] badges (AC-01)
- [ ] Badge styling matches tool registry panel (AC-01)

---

## TASK-006-06-05: Add client configuration guidance text for Gemini CLI and Claude Code setup

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- Add a collapsible "Client Setup Guide" section (using `Expander`) below the Configuration panel or as a subsection within it
- Include guidance text for Gemini CLI: "Add to your Gemini CLI MCP configuration: `{\"url\": \"http://localhost:{port}/sse\"}`"
- Include guidance for Claude Code: "Add to your Claude Code MCP settings: `{\"url\": \"http://localhost:{port}/sse\"}`"
- Bind the port in the guidance text to `{Binding ServerPort}` so it updates dynamically
- Add a "Copy URL" button next to the SSE endpoint URL that copies `http://localhost:{port}/sse` to the clipboard using `Clipboard.SetText()`
- Keep the section collapsed by default to avoid visual clutter (per FR-006-048 priority "Could")

### How to verify
- [ ] Client configuration guidance is displayed (AC-06)
- [ ] Guidance includes correct port number (AC-06)
- [ ] Copy URL button works correctly (AC-06)
- [ ] URLs update when port changes (AC-07)

---

## Dependency Graph
```
TASK-006-06-01 (XAML Configuration Panel)
       (independent)

TASK-006-06-02 (Computed Endpoint URLs)
       (independent)

TASK-006-06-03 (GUI Proxy Status Display)
       (independent)

TASK-006-06-04 (Prerequisite Legend)
       (independent)

TASK-006-06-05 (Client Setup Guide)
       (independent)

All five tasks are independent and can be developed in parallel.
They each contribute to different sections of the Configuration panel.
```
