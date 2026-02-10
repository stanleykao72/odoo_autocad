# US-008-05: Window Default Size

## User Story
**As a** System Admin,
**I want to** have the application start centered on screen at a reasonable default size,
**So that** the window is immediately usable without manual resizing.

## Parent Feature
- **FR**: [FR-008-ui-framework](../FR-008-ui-framework/FR-008-ui-framework.md)
- **Priority**: P1

## Acceptance Criteria
- [x] AC-01: The main window starts with a default size of 1280x720 pixels.
- [x] AC-02: The main window enforces a minimum size of 900x600 pixels that cannot be overridden at runtime.
- [x] AC-03: The main window is centered on screen using `WindowStartupLocation="CenterScreen"`.
- [x] AC-04: The main window uses a two-column Grid layout: fixed-width sidebar (250px left) and flexible content area (remaining width, `*`).
- [x] AC-05: The content area contains three rows: header (Auto height), content Frame (`*` fills remaining), and status bar (Auto height).
- [x] AC-06: The window background uses `BackgroundBrush` from the application resource dictionary.
- [x] AC-07: The sidebar width remains fixed at 250px while the content area fills the remaining space on resize.
- [x] AC-08: The content Frame expands to fill all available vertical space between the header and status bar.
- [x] AC-09: The window is resizable with proper content reflow and minimum size constraints enforced.

## Related Functional Requirements
| ID | Description | Priority |
|----|-------------|----------|
| FR-008-001 | Two-column Grid layout: fixed sidebar (250px) and flexible content area | Must |
| FR-008-002 | Content area with three rows: header, content Frame, status bar | Must |
| FR-008-003 | Default size 1280x720, minimum size 900x600 | Must |
| FR-008-004 | WindowStartupLocation="CenterScreen" | Must |
| FR-008-005 | Background from BackgroundBrush resource | Must |
| FR-008-036 | Sidebar fixed at 250px, content fills remaining space | Must |
| FR-008-037 | Content Frame expands vertically between header and status bar | Must |
| FR-008-038 | Resizable window with content reflow and minimum size constraints | Must |

## Implementation Tasks
| Task ID | Description | Target File | Estimate |
|---------|-------------|-------------|----------|
| TASK-008-05-01 | Define MainWindow XAML with Height, Width, MinHeight, MinWidth, WindowStartupLocation, and Background properties | `Views/MainWindow.xaml` | S |
| TASK-008-05-02 | Define two-column Grid with fixed 250px sidebar column and star-sized content column | `Views/MainWindow.xaml` | S |
| TASK-008-05-03 | Define three-row content area: Auto header, star content Frame, Auto status bar | `Views/MainWindow.xaml` | S |
| TASK-008-05-04 | Define BackgroundBrush resource in application resource dictionary | `Themes/Colors.xaml` | S |
| TASK-008-05-05 | Set window Title to "Odoo AutoCAD Integration" and Icon to application icon | `Views/MainWindow.xaml` | S |

## Dependencies
- Depends on: None (this is a standalone window layout concern)
- Blocks: None directly, but the layout structure is consumed by US-008-01, US-008-03, US-008-04, US-008-08

## Notes
- The Python implementation uses `setup_window_geometry()` with min 1000x700, default 80% of screen (max 1200x800), centered. The C# port uses fixed defaults (1280x720, min 900x600) defined declaratively in XAML, which is simpler and more predictable.
- Validation rule VR-008-008 requires that the minimum size (900x600) is enforced by `MinHeight`/`MinWidth` properties and must not be overridable at runtime.
- The `BackgroundBrush` is defined as `#F5F5F5` in the resource dictionary (see FR-008-032).
- WPF handles resize reflow automatically through the Grid layout system with proper column/row definitions.
