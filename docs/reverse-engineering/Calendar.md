# Calendar / CalendarDate — Legacy Field Configuration

## Type Identifiers

| Property | Value |
|----------|-------|
| `cluster_type` | `Calendar` or `CalendarDate` |
| XML `<type>` | `Calendar` / `CalendarDate` |
| Comment line 1 | `Calendar` |
| Parameter class | `DateClusterParameter` |
| DLL class count | `DateClusterParameter` (3 hits) |

## Sources

| Source | Evidence |
|--------|----------|
| **LibConMas.dll** | `DateClusterParameter` class (3 hits), `CalendarDate` (56 hits in broad search) |
| **ConMas.iReporter.UserControls.dll** | Calendar/Date WPF picker controls |
| **ConMasClient.exe** | `CalendarDateForm`, `CalendarDateDialog`, `CalendarDateTimeForm` |
| **LocalizableStrings.xml** | `CalendarLibrary.Monday` (and other day names) — confirms calendar library |
| **Cimtops.R2Cluster.dll** | `CalendarDate` referenced in type list |

## Configuration Keys

| Key | Type | Default | Values | Description | Evidence |
|-----|------|---------|--------|-------------|----------|
| `Required` | bool | 0 | 0/1 | Whether date selection is mandatory | Common parameter |
| `DateFormat` | string | (system) | Date format string | Display format for the selected date | LibConMas (188 hits, strongest of all params) |
| `TimeFormat` | string | (system) | Time format string | Display format for time (if calendar includes time) | LibConMas (64 hits) |
| `ReadOnly` | bool | 0 | 0/1 | Whether date is read-only | ClusterParameter base |
| `Hidden` | bool | 0 | 0/1 | Whether field is hidden | UserControls |
| `DefaultValue` | string | (empty) | Date string | Default date | Common parameter |

### Hypothetical Keys

| Key | Type | Description | Evidence |
|-----|------|-------------|----------|
| `CalendarType` | enum | Gregorian/Japanese/etc. | Naming pattern only |
| `ShowWeekNumber` | bool | Show week numbers in calendar | Common calendar feature |
| `FirstDayOfWeek` | enum | Sunday/Monday/etc. | Common calendar feature |
| `MinDate` | string | Minimum selectable date | Pattern from InputNumeric MinValue/MaxValue |
| `MaxDate` | string | Maximum selectable date | Pattern from InputNumeric MinValue/MaxValue |

## Designer Property Pages (ConMasClient.exe)

- `CalendarDateForm` — Calendar date property form
- `CalendarDateDialog` — Calendar date picker dialog
- `CalendarDateTimeForm` — Combined calendar + time form
- `DateDetermination` / `DateDeterminnation` — Date determination settings (8 hits)

### Date-related forms:
- `RegistrationDateForm` — Registration date (system)
- `LatestUpdateDateForm` — Latest update date (system)
- `ExpiryDateForm` — Expiry date (system)

## Runtime Behavior

The legacy runtime uses a WPF `DatePicker` or custom calendar control for date selection.

### Day Names (from LocalizableStrings.xml)
`CalendarLibrary.Monday`, etc. — Calendar library has localized day names in Japanese, English, Chinese.

## Related System Types

| Field Type | Description | Evidence |
|------------|-------------|----------|
| `LatestUpdateDate` | Auto-populated with last update timestamp | `LatestUpdateDateClusterParameter` |
| `RegistrationDate` | Auto-populated with creation timestamp | `RegistrationDateClusterParameter` |
| `ExpiryDate` | Auto-populated with expiry date | `ExpiryDateForm` |
| `Date` | Simple date field (no calendar UI?) | `DateClusterParameter` class |

## Storage

- **Database:** `def_cluster.input_parameter` column
- **XML:** `<inputParameters>...</inputParameters>` within `<cluster>` element
- **Memory:** `DateClusterParameter` class

## Confidence Summary

| Item | Confidence |
|------|------------|
| DateFormat parameter | ★★★★★ |
| TimeFormat parameter | ★★★★☆ |
| CalendarDate class exists | ★★★★★ |
| Calendar designer dialogs exist | ★★★★☆ |
| Runtime DatePicker control | ★★★★☆ |
