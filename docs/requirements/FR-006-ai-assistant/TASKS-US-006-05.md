# TASKS: US-006-05 — View Tool Registry

> **Parent US**: [US-006-05](US-006-05-view-tool-registry.md)
> **Parent FR**: [FR-006](FR-006-ai-assistant.md)
> **Priority**: P3
> **Tasks**: 6 | **Effort**: 3S + 3M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] FR-008 (UI Framework - WPF/MVVM infrastructure must be in place)
- [ ] `MCPToolRegistry` exists in `OdooAutoCAD.MCP/Tools/MCPToolRegistry.cs` with `GetTools()` method
- [ ] `MCPTool` model exists in `OdooAutoCAD.MCP/Protocol/MCPModels.cs` with Name, Description, InputSchema

## Acceptance Criteria
- [ ] AC-01: The page displays a complete list of all registered MCP tools with name and description
- [ ] AC-02: Tools are grouped by category (Connection, AutoCAD Drawing, Data Integration, Layout, NLP)
- [ ] AC-03: Each tool entry shows its input parameters schema including parameter names, types, and required/optional status
- [ ] AC-04: Each tool entry indicates prerequisite connections with badges: [A] for AutoCAD-required, [O] for Odoo-required, or both [A][O]
- [ ] AC-05: The tool registry supports dynamic registration via MCPToolRegistry.RegisterTool()
- [ ] AC-06: The total tool count is displayed in the server status panel

---

## TASK-006-05-01: Create Registered Tools panel XAML with grouped ListView and category headers

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-006-05-04 |

### What to do
- Add a "Registered Tools" section in the left-center area of `MCPAssistantPage.xaml` using a `Border` with header TextBlock
- Use an `ItemsControl` or `ListView` with `ItemsSource="{Binding RegisteredTools}"` and `GroupStyle` for category grouping
- Define a `DataTemplate` for each tool entry showing: tool name (bold TextBlock), description (secondary TextBlock), and a placeholder for prerequisite badges
- Add a `CollectionViewSource` in page resources with `GroupDescriptions` on the `Category` property: `<CollectionViewSource.GroupDescriptions><PropertyGroupDescription PropertyName="Category"/></CollectionViewSource.GroupDescriptions>`
- Define `GroupStyle` with a header template showing category name and tool count in parentheses (e.g., "Category: Connection (4)")
- Make the list scrollable with `ScrollViewer.VerticalScrollBarVisibility="Auto"`

### How to verify
- [ ] Complete list of registered tools is displayed with name and description (AC-01)
- [ ] Tools are visually grouped by category with headers (AC-02)

---

## TASK-006-05-02: Implement MCPToolViewModel with display properties

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-006-05-03 |

### What to do
- Create `MCPToolViewModel` class (can be a nested class or separate file) with properties: `string Name`, `string Description`, `string Category`, `bool RequiresAutoCAD`, `bool RequiresOdoo`, `MCPInputSchema InputSchema`
- Add a computed property `string PrerequisiteBadges` that returns "[A]", "[O]", "[A][O]", or "" based on the boolean prerequisite flags
- Add a computed property `string ParametersSummary` that formats the InputSchema properties as a readable string (e.g., "drawing_path: string (optional), use_current_drawing: boolean (optional)")
- Add `bool IsParametersExpanded` observable property (default false) for expandable parameter details
- Create a mapping method `static MCPToolViewModel FromMCPTool(MCPTool tool)` to convert from the protocol model, setting Category, RequiresAutoCAD, RequiresOdoo based on tool name mappings

### How to verify
- [ ] MCPToolViewModel contains all required display properties (AC-01, AC-02, AC-04)
- [ ] PrerequisiteBadges correctly reflects tool requirements (AC-04)
- [ ] ParametersSummary formats input schema readably (AC-03)

---

## TASK-006-05-03: Populate RegisteredTools ObservableCollection from MCPToolRegistry.GetTools()

| Field | Value |
|-------|-------|
| Target | `ViewModels/MCPViewModel.cs` |
| Estimate | S |
| Depends On | TASK-006-05-02 |
| Blocks | None |

### What to do
- Add `ObservableCollection<MCPToolViewModel> RegisteredTools` property to `MCPViewModel`
- In `MCPViewModel` constructor or initialization, call `_toolRegistry.GetTools()` and convert each `MCPTool` to `MCPToolViewModel` using the mapping method
- Populate `RegisteredTools` collection and set `RegisteredToolCount = RegisteredTools.Count`
- Subscribe to `MCPToolRegistry.ToolRegistered` event (if implemented) to dynamically add new tools to the collection
- When a new tool is registered dynamically, create a new `MCPToolViewModel`, add it to the `ObservableCollection`, and increment `RegisteredToolCount`
- Use `Dispatcher.InvokeAsync` for thread-safe collection updates if events arrive from background threads

### How to verify
- [ ] RegisteredTools collection contains all tools from the registry (AC-01)
- [ ] RegisteredToolCount matches the collection count (AC-06)
- [ ] Dynamic tool registration updates the list automatically (AC-05)

---

## TASK-006-05-04: Add GroupStyle to ItemsControl for category-based grouping with CollectionViewSource

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | M |
| Depends On | TASK-006-05-01 |
| Blocks | None |

### What to do
- Define a `CollectionViewSource` in the page or section resources: `<CollectionViewSource x:Key="GroupedTools" Source="{Binding RegisteredTools}">`
- Add `GroupDescriptions`: `<PropertyGroupDescription PropertyName="Category"/>`
- Set the `ItemsControl.ItemsSource` to `{Binding Source={StaticResource GroupedTools}}`
- Create a `GroupStyle.HeaderTemplate` `DataTemplate` with a `TextBlock` showing `"{Binding Name} ({Binding ItemCount})"` styled as a section header (bold, larger font, with bottom border)
- Style the group headers with category-specific background colors or icons (optional enhancement)
- Ensure categories appear in logical order: Connection, AutoCAD Drawing, Data Integration, Layout, NLP (use a custom `IComparer` on the `PropertyGroupDescription` if needed)

### How to verify
- [ ] Tools are grouped under category headers (AC-02)
- [ ] Each category header shows the name and tool count (AC-02)
- [ ] Groups appear in logical category order (AC-02)

---

## TASK-006-05-05: Display prerequisite badges next to tool names using DataTemplate

| Field | Value |
|-------|-------|
| Target | `Views/Pages/MCPAssistantPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the tool entry `DataTemplate`, add a `StackPanel Orientation="Horizontal"` containing the tool name TextBlock followed by badge elements
- Add a `Border` with text "[A]" (blue/cyan background, white text, small CornerRadius) visible when `RequiresAutoCAD` is true, using `BooleanToVisibilityConverter`
- Add a `Border` with text "[O]" (purple/orange background, white text, small CornerRadius) visible when `RequiresOdoo` is true
- Add small `Margin` between badges for spacing
- Apply `ToolTip` to each badge: "[A]" shows "Requires AutoCAD connection", "[O]" shows "Requires Odoo connection"
- Below the name+badges row, add the description TextBlock with secondary text style
- Optionally add an expandable section (Expander or toggle) for the InputSchema parameter details bound to `ParametersSummary` and `IsParametersExpanded`

### How to verify
- [ ] [A] badge appears for AutoCAD-required tools (AC-04)
- [ ] [O] badge appears for Odoo-required tools (AC-04)
- [ ] [A][O] badges appear together for tools requiring both (AC-04)
- [ ] Parameter schema is viewable per tool entry (AC-03)

---

## TASK-006-05-06: Ensure MCPToolRegistry.RegisterTool() triggers collection update via event notification

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.MCP/Tools/MCPToolRegistry.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Make the `RegisterTool()` method `public` (currently private) to support dynamic registration per FR-006-013
- Add an `event Action<MCPTool>? ToolRegistered` event to `MCPToolRegistry`
- In `RegisterTool()`, after adding the tool to `_toolDefinitions` and `_toolExecutors`, fire `ToolRegistered?.Invoke(tool)`
- Add validation: if a tool with the same name already exists, log a warning and either replace (update) or throw `InvalidOperationException` (VR-006-008 -- all tool names must be unique)
- Add a `public int ToolCount => _toolDefinitions.Count` convenience property
- Add a `public bool HasTool(string name) => _toolDefinitions.ContainsKey(name)` method
- Add `Category`, `RequiresAutoCAD`, `RequiresOdoo` properties to the `MCPTool` model class in `MCPModels.cs` (add `[JsonIgnore]` attributes since these are metadata not sent over the wire)

### How to verify
- [ ] Dynamic registration via RegisterTool() is supported (AC-05)
- [ ] ToolRegistered event fires on new registration (AC-05)
- [ ] Duplicate tool names are rejected or warned (AC-05)
- [ ] Tool count is available via ToolCount property (AC-06)

---

## Dependency Graph
```
TASK-006-05-01 (XAML Tools Panel)
       │
       └──▶ TASK-006-05-04 (GroupStyle + CollectionViewSource)

TASK-006-05-02 (MCPToolViewModel)
       │
       └──▶ TASK-006-05-03 (Populate Collection)

TASK-006-05-05 (Prerequisite Badges XAML)
       (independent)

TASK-006-05-06 (MCPToolRegistry Dynamic Registration)
       (independent)

Tasks 01, 02, 05, and 06 are independent and can be developed in parallel.
Task 04 depends on 01 (XAML panel must exist for GroupStyle).
Task 03 depends on 02 (MCPToolViewModel must exist to populate collection).
```
