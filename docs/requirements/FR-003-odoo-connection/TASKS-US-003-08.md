# TASKS: US-003-08 — View Last Sync Time

> **Parent US**: [US-003-08](US-003-08-view-last-sync-time.md)
> **Parent FR**: [FR-003](FR-003-odoo-connection.md)
> **Priority**: P1
> **Tasks**: 6 | **Effort**: 4S + 2M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-003-05 (product sync) must be completed so sync operations exist to timestamp

## Acceptance Criteria
- [ ] AC-01: The last synchronization timestamp is displayed on the Odoo Integration Page
- [ ] AC-02: The timestamp is retrieved via `GetLastSyncTime()` from the IOdooService
- [ ] AC-03: The timestamp is formatted in the user's local timezone
- [ ] AC-04: The timestamp updates automatically after each successful sync operation
- [ ] AC-05: If no sync has ever been performed, the display shows "Never synced" or equivalent
- [ ] AC-06: The sync timestamp is visible in both the Connection Status panel and the Sync Controls panel

---

## TASK-003-08-01: Add last sync timestamp labels to XAML panels

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Views/Pages/OdooConnectionPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Connection Status panel (created in US-003-03), add a row:
  - "Last Sync:" label + `TextBlock` bound to `{Binding LastSyncTime, Converter={StaticResource DateTimeToStringConverter}}`
- In the Sync Controls panel (right side of the page per wireframe), add:
  - "Last Sync:" label + `TextBlock` with the same binding
  - "Records:" label + `TextBlock` bound to `{Binding LastSyncRecordsSummary}` (e.g., "1,247 products")
  - "Status:" label + `TextBlock` bound to `{Binding LastSyncStatus}` (e.g., "Success" / "Failed")
- Both locations use the same ViewModel property with the `DateTimeToStringConverter` for local timezone formatting
- Use `StringFormat` or a converter to handle the "Never synced" fallback when value is null

### How to verify
- [ ] Last sync timestamp is displayed on the page (AC-01)
- [ ] Timestamp is visible in both Connection Status and Sync Controls panels (AC-06)

---

## TASK-003-08-02: Add LastSyncTime property to ViewModel

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-003-08-05 |

### What to do
- Add observable properties using `[ObservableProperty]`:
  - `private DateTime? _lastSyncTime;`
  - `private string _lastSyncStatus = string.Empty;`
  - `private string _lastSyncRecordsSummary = string.Empty;`
- Implement `partial void OnLastSyncTimeChanged(DateTime? value)`:
  - Trigger `OnPropertyChanged(nameof(LastSyncTimeFormatted))` if using a computed property for display
- Add a computed property for formatted display:
  ```csharp
  public string LastSyncTimeFormatted => LastSyncTime?.ToLocalTime().ToString("yyyy-MM-dd HH:mm") ?? "Never synced";
  ```
- Initialize `LastSyncTime` from `_odooService.GetLastSyncTime()` in the constructor or `OnNavigatedTo`

### How to verify
- [ ] LastSyncTime property exists with nullable DateTime type (AC-02)
- [ ] "Never synced" displayed when no sync has occurred (AC-05)

---

## TASK-003-08-03: Implement GetLastSyncTime in OdooService

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Core/Odoo/OdooService.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- The existing `GetLastSyncTime()` method returns `_lastSyncTime` which is an in-memory field:
  ```csharp
  public DateTime? GetLastSyncTime() => _lastSyncTime;
  ```
- This works for in-session tracking but does not survive application restarts
- Modify `GetLastSyncTime()` to first check the in-memory value, then fall back to database:
  ```csharp
  public DateTime? GetLastSyncTime()
  {
      if (_lastSyncTime.HasValue) return _lastSyncTime;
      // Delegate to persistence layer if needed
      return _persistedLastSyncTime;
  }
  ```
- Update `_lastSyncTime` at the end of `ImportToBOQAsync()`, `SyncToOdooAsync()`, and in the ViewModel after `GetProductsAsync()`
- Add a `SetLastSyncTime(DateTime time)` internal method for explicit updates

### How to verify
- [ ] GetLastSyncTime returns the stored last sync timestamp (AC-02)
- [ ] Returns null when no sync has been performed (AC-05)

---

## TASK-003-08-04: Persist last sync time to local database

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.Data/Context/AppDbContext.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Use the existing `SyncLog` entity in `AppDbContext` to persist sync timestamps:
  ```csharp
  public class SyncLog
  {
      public int Id { get; set; }
      public string SyncType { get; set; }  // "product", "boq", "project", "parameters"
      public DateTime SyncTime { get; set; }
      public bool Success { get; set; }
      public int RecordsProcessed { get; set; }
      ...
  }
  ```
- Create a `SyncLogService` or repository class:
  - `Task SaveSyncLogAsync(SyncLog log)` -- insert new sync log entry
  - `Task<DateTime?> GetLastSyncTimeAsync(string? syncType = null)` -- query the most recent successful sync time, optionally filtered by type
  - `Task<IReadOnlyList<SyncLog>> GetSyncHistoryAsync(int limit = 10)` -- retrieve recent sync history
- Query using EF Core: `await _dbContext.SyncLogs.Where(l => l.Success).OrderByDescending(l => l.SyncTime).FirstOrDefaultAsync()`
- Load the persisted last sync time during ViewModel initialization to populate the display immediately

### How to verify
- [ ] Last sync time survives application restarts via database persistence (AC-01)
- [ ] Per-type sync timestamps are available for detailed history (AC-04)

---

## TASK-003-08-05: Update LastSyncTime after each successful sync operation

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/ViewModels/OdooConnectionViewModel.cs` |
| Estimate | S |
| Depends On | TASK-003-08-02 |
| Blocks | None |

### What to do
- At the end of `SyncProductsAsync()` (after successful product fetch):
  ```csharp
  LastSyncTime = DateTime.Now;
  LastSyncStatus = "Success";
  LastSyncRecordsSummary = $"{TotalProductCount:N0} products";
  ```
- At the end of `SyncToOdooAsync()` (after BOQ or parameter sync):
  ```csharp
  LastSyncTime = DateTime.Now;
  LastSyncStatus = result.Success ? "Success" : (result.RecordsFailed > 0 ? "Partial" : "Failed");
  ```
- Optionally persist the sync log entry via `SyncLogService.SaveSyncLogAsync()`
- The timestamp should update in real-time on the UI via property change notification
- On partial failure, still update `LastSyncTime` but set `LastSyncStatus = "Partial"`

### How to verify
- [ ] Timestamp updates automatically after each successful sync (AC-04)
- [ ] Status reflects success, partial failure, or failure (AC-04)

---

## TASK-003-08-06: Create DateTime-to-string converter with "Never synced" fallback

| Field | Value |
|-------|-------|
| Target | `src/OdooAutoCAD.App/Converters/DateTimeToStringConverter.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | None |

### What to do
- Create `DateTimeToStringConverter` class implementing `IValueConverter` in namespace `OdooAutoCAD.App.Converters`
- `Convert` method:
  - If value is `null` or `DependencyProperty.UnsetValue`, return `"Never synced"`
  - If value is `DateTime`, convert to local timezone and format as `"yyyy-MM-dd HH:mm"` matching the wireframe format
  - If parameter is provided (e.g., `ConverterParameter="relative"`), return relative time like "5 minutes ago"
- `ConvertBack` method: throw `NotSupportedException` (one-way binding only)
- Register as a `StaticResource` in `App.xaml` or `OdooConnectionPage.xaml`:
  ```xml
  <converters:DateTimeToStringConverter x:Key="DateTimeToStringConverter" />
  ```
- Handle edge case: UTC DateTime stored in DB should be converted to `ToLocalTime()` before formatting

### How to verify
- [ ] Timestamp is formatted in user's local timezone (AC-03)
- [ ] "Never synced" shown when no sync has occurred (AC-05)

---

## Dependency Graph
```
TASK-003-08-02 (ViewModel LastSyncTime property)
    |
    +---> TASK-003-08-05 (Update after sync operations)

TASK-003-08-01 (XAML timestamp labels)

TASK-003-08-03 (OdooService GetLastSyncTime)

TASK-003-08-04 (Database persistence)

TASK-003-08-06 (DateTime converter)
```
