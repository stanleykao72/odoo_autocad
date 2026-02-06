# TASKS: US-007-10 — Clear Cache

> **Parent US**: [US-007-10](US-007-10-clear-cache.md)
> **Parent FR**: [FR-007](FR-007-settings.md)
> **Priority**: P2
> **Tasks**: 8 | **Effort**: 5S + 3M + 0L
> **Status**: Not Started

## Prerequisites
- [ ] US-007-01 (SettingsPage XAML and SettingsViewModel must exist with TabControl structure)
- [ ] `AppDbContext` with `BOQCache`, `ProductMappings`, `SyncLogs` DbSets available
- [ ] `IOdooService` interface with cache invalidation method available (or to be added)

## Acceptance Criteria
- [ ] AC-01: Settings Page displays a "Clear Cache" button in the Advanced tab under Data Management section
- [ ] AC-02: Clicking "Clear Cache" shows a confirmation dialog with message: "This will remove {n} cached BOQ entries and {m} product mappings. This action cannot be undone. Continue?"
- [ ] AC-03: The confirmation dialog displays the actual record counts for BOQCache and ProductMapping tables
- [ ] AC-04: Confirming the dialog removes all records from `BOQCache` and `ProductMapping` database tables
- [ ] AC-05: Confirming the dialog also resets relevant `SyncLog` entries
- [ ] AC-06: Confirming the dialog invalidates any in-memory caches held by `IOdooService`
- [ ] AC-07: After successful cache clear, a success message with deleted record counts is displayed
- [ ] AC-08: Cancelling the confirmation dialog takes no action and returns to the settings page
- [ ] AC-09: If cache clear fails, an error dialog is shown: "Failed to clear cache. Error: {details}"
- [ ] AC-10: Cache clear operation is logged to the `SyncLog` table with SyncType "settings_change"

---

## TASK-007-10-01: Add "Clear Cache" button to Data Management section of Advanced tab in XAML

| Field | Value |
|-------|-------|
| Target | `Views/Pages/SettingsPage.xaml` |
| Estimate | S |
| Depends On | None |
| Blocks | None |

### What to do
- In the Advanced tab, add a "Data Management" section group (below the Logging section from US-007-09)
- Add a `Button` with `Content="Clear Cache"` bound to `{Binding ClearCacheCommand}`
- Add a `TextBlock` for result feedback (success/error message) bound to a feedback property, with `Visibility` collapsed when empty
- Style the button with a warning appearance (e.g., orange/red accent) to indicate destructive operation
- Position alongside Export and Import buttons (which will be added in US-007-11 and US-007-12)

### How to verify
- [ ] "Clear Cache" button renders in the Data Management section of the Advanced tab (AC-01)
- [ ] Feedback text area is present but hidden when no message (AC-07, AC-09)

---

## TASK-007-10-02: Implement ClearCacheCommand as IAsyncRelayCommand in ViewModel

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | M |
| Depends On | None |
| Blocks | TASK-007-10-04, TASK-007-10-05, TASK-007-10-07, TASK-007-10-08 |

### What to do
- Add `IAsyncRelayCommand ClearCacheCommand` initialized as `new AsyncRelayCommand(ClearCacheAsync)`
- Implement `private async Task ClearCacheAsync()` skeleton that:
  1. Queries record counts (TASK-007-10-03)
  2. Shows confirmation dialog (TASK-007-10-04)
  3. On confirm: executes deletion (TASK-007-10-05), invalidates caches (TASK-007-10-06), shows feedback (TASK-007-10-07)
  4. On cancel: returns without action (AC-08)
  5. Wraps in try/catch for error handling (TASK-007-10-07)
- Add `[ObservableProperty] string _cacheOperationMessage` (default empty) for success/error feedback

### How to verify
- [ ] `ClearCacheCommand` is callable from XAML binding (AC-01)
- [ ] Command orchestrates the full clear cache workflow (AC-02 through AC-10)

---

## TASK-007-10-03: Query BOQCache and ProductMapping record counts for confirmation dialog

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | None |
| Blocks | TASK-007-10-04 |

### What to do
- In `ClearCacheAsync()`, query the database for actual record counts:
  - `int boqCount = await _dbContext.BOQCache.CountAsync()`
  - `int mappingCount = await _dbContext.ProductMappings.CountAsync()`
- Store these counts for use in the confirmation dialog message and in the success feedback
- If both counts are zero, optionally skip the confirmation and show "Cache is already empty" message

### How to verify
- [ ] Record counts are queried from the actual database tables (AC-03)
- [ ] Counts reflect the real number of records, not estimates (AC-03)

---

## TASK-007-10-04: Implement confirmation dialog with record counts and Cancel/Continue buttons

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-10-02, TASK-007-10-03 |
| Blocks | TASK-007-10-05 |

### What to do
- Show a `MessageBox` or custom dialog with:
  - Title: "Clear Cache"
  - Message: `$"This will remove {boqCount} cached BOQ entries and {mappingCount} product mappings. This action cannot be undone. Continue?"`
  - Buttons: "Continue" (Yes) and "Cancel" (No)
  - Icon: Warning (exclamation mark)
- If the user clicks Cancel, return from `ClearCacheAsync()` without performing any deletion
- Use `IDialogService` abstraction if available for testability; otherwise use `MessageBox.Show()` with `MessageBoxButton.YesNo`

### How to verify
- [ ] Confirmation dialog shows the correct message with record counts (AC-02, AC-03)
- [ ] Cancelling the dialog takes no action (AC-08)

---

## TASK-007-10-05: Execute RemoveRange on BOQCache and ProductMapping tables via AppDbContext

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | M |
| Depends On | TASK-007-10-04 |
| Blocks | TASK-007-10-06, TASK-007-10-07 |

### What to do
- After user confirms, execute the deletion within a database transaction:
  ```csharp
  using var transaction = await _dbContext.Database.BeginTransactionAsync();
  try
  {
      _dbContext.BOQCache.RemoveRange(_dbContext.BOQCache);
      _dbContext.ProductMappings.RemoveRange(_dbContext.ProductMappings);
      // Reset relevant SyncLog entries (AC-05)
      await _dbContext.SaveChangesAsync();
      await transaction.CommitAsync();
  }
  catch
  {
      await transaction.RollbackAsync();
      throw;
  }
  ```
- Use `RemoveRange()` to efficiently delete all records
- Include resetting relevant `SyncLog` entries: add entries marking the sync as invalidated, or delete BOQ/product-related sync logs as appropriate

### How to verify
- [ ] After confirm, `BOQCache` table is empty (AC-04)
- [ ] After confirm, `ProductMappings` table is empty (AC-04)
- [ ] Relevant `SyncLog` entries are reset (AC-05)
- [ ] Deletion is transactional -- if any part fails, all changes are rolled back (AC-09)

---

## TASK-007-10-06: Invalidate in-memory caches in IOdooService after database clear

| Field | Value |
|-------|-------|
| Target | `OdooAutoCAD.Core/Odoo/IOdooService.cs` and implementation |
| Estimate | S |
| Depends On | TASK-007-10-05 |
| Blocks | None |

### What to do
- Add `void InvalidateCache()` method to `IOdooService` interface (if not already present)
- Implement in the concrete `OdooService`: clear any in-memory dictionaries, lists, or cached product data
- Call `_odooService.InvalidateCache()` in `ClearCacheAsync()` after successful database deletion
- This ensures the application does not serve stale data from memory after a cache clear

### How to verify
- [ ] After cache clear, `IOdooService` no longer returns previously cached data (AC-06)
- [ ] Subsequent data requests fetch fresh data from the Odoo server (AC-06)

---

## TASK-007-10-07: Add success/error result feedback after cache clear operation

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-10-02, TASK-007-10-05 |
| Blocks | None |

### What to do
- On successful deletion: set `CacheOperationMessage = $"Cache cleared successfully. Removed {boqCount} BOQ entries and {mappingCount} product mappings."`
- On failure (caught `DbUpdateException` or other exception): set `CacheOperationMessage = $"Failed to clear cache. Error: {ex.InnerException?.Message ?? ex.Message}"` and optionally show a `MessageBox` error dialog
- Clear the feedback message after a timeout (e.g., 10 seconds) or on the next user interaction
- Style success messages in green and error messages in red in the XAML

### How to verify
- [ ] Success message displays with correct deleted record counts (AC-07)
- [ ] Error message displays with exception details on failure (AC-09)

---

## TASK-007-10-08: Log cache clear operation to SyncLog table

| Field | Value |
|-------|-------|
| Target | `ViewModels/SettingsViewModel.cs` |
| Estimate | S |
| Depends On | TASK-007-10-02 |
| Blocks | None |

### What to do
- After successful cache clear, add a `SyncLog` entry:
  - `SyncType = "settings_change"`
  - `Direction = "upload"`
  - `RecordsProcessed = boqCount + mappingCount`
  - `Success = true`
  - `Details` = JSON string: `{ "action": "clear_cache", "boq_deleted": {n}, "product_mappings_deleted": {m} }`
  - `SyncTime = DateTime.UtcNow`
- If cache clear fails, log with `Success = false` and `ErrorMessage` set to the exception message
- Call `await _dbContext.SaveChangesAsync()` to persist the log entry

### How to verify
- [ ] After successful clear, `SyncLog` contains entry with `SyncType = "settings_change"` and cache details (AC-10)
- [ ] Failed clear also logs with `Success = false` (AC-10)

---

## Dependency Graph
```
TASK-007-10-01 (XAML Button)         ── independent
TASK-007-10-03 (Query Counts)        ── independent

TASK-007-10-02 (ClearCacheCommand)
       │
       ├──▶ TASK-007-10-04 (Confirmation Dialog) ◀── TASK-007-10-03
       │         │
       │         └──▶ TASK-007-10-05 (Execute Deletion)
       │                   │
       │                   ├──▶ TASK-007-10-06 (Invalidate In-Memory)
       │                   └──▶ TASK-007-10-07 (Feedback)
       │
       └──▶ TASK-007-10-08 (SyncLog Audit)

Tasks 01, 02, and 03 can start in parallel.
```
