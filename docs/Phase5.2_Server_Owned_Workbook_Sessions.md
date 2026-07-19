# Phase 5.2 — Eliminate Re-Upload Requirement (Server-Owned Workbook Sessions)

**Status:** Complete ✅  
**Date:** July 19, 2026  
**Build:** 0 errors, 79 warnings (pre-existing)

---

## Objective

Remove the requirement for the browser to remember the uploaded Excel filename (`sourceFileName`) or ask the user to upload the workbook again.

The user experience now matches the legacy ConMas Designer:

```
Upload Excel
      ↓
Preview appears
      ↓
Edit fields
      ↓
Export Excel
      ↓
Download edited workbook
```

**No re-upload. No browser dependency. No "No source workbook available" messages.**

---

## Legacy ConMas Architecture (Reverse-Engineered from Docs)

The ConMas Designer always owned the workbook server-side:

```
Upload Flow:
  Designer → Base64-encodes entire XLSX in <definitionFile><value>
          → Sends as XML via HTTP POST
          → Server decodes to BLOB in def_top_file column
          → Returns defTopId (integer PK)
          → Browser only tracks defTopId

Save Flow:
  Browser sends defTopId → Server loads XLSX from DB BLOB
                         → Re-renders via Syncfusion
                         → Overlays field values from rep_cluster
```

Key insight: **The server always owned the workbook.** The browser never tracked filenames or physical paths.

---

## What We Built

### New: `SessionWorkbookStore` — Server-Side Session Storage

| File | `ExcelAPI/Application/SessionWorkbookStore.cs` |
|------|-------------------------------------------------|
| Interface | `ISessionWorkbookStore` |
| Storage | `TempWorkbooks/{sessionId}/original.xlsx` |
| Session lifetime | 24 hours |
| Thread safety | `ConcurrentDictionary<string, SessionInfo>` |
| Disk persistence | Falls back to disk if in-memory cache lost |

### Changed: `WorkbookDefinition` Model

| Change | Details |
|--------|---------|
| Added | `string SessionId` — primary key for session-based save |
| Deprecated | `string SourceFileName` — `[Obsolete("Use SessionId instead")]` |

### Changed: Upload Endpoints

| Endpoint | Old Behavior | New Behavior |
|----------|-------------|--------------|
| `POST /api/form/from-excel` | Saved to `Forms/`, returned `templateId` | Saves to session store, returns `sessionId` |
| `POST /api/form/upload-preview` | Deleted XLSX after preview | Saves to session store, returns `sessionId` |
| `POST /api/form/upload-excel` | Saved to `Forms/`, returned `templateId` + `workbookDownloadUrl` | Saves to session store, returns `sessionId` |

### Changed: `POST /api/form/save-edited`

| Aspect | Before | After |
|--------|--------|-------|
| Source resolution | `SourceFileName` → `Forms/{filename}` | `SessionId` → `SessionWorkbookStore.ResolveWorkbookPath()` |
| Error on session expiry | N/A (no sessions) | HTTP 410: "The editing session has expired" |
| Fallback | None | `SourceFileName` (deprecated) for backward compat |

### Changed: Frontend (`page.tsx`)

| Change | Before | After |
|--------|--------|-------|
| State variable | `const [sourceFileName, setSourceFileName]` | `const [sessionId, setSessionId]` |
| `handleUpload()` | Stored `result.templateId` as filename | Stored `result.sessionId` |
| `handleUploadExcel()` | Stored `result.data.templateId` as filename | Stored `result.data.sessionId` |
| `runtimeFormToWorkbookDefinition()` | Accepted `source: string`, set `sourceFileName` | Accepts `sid: string`, sets `sessionId` + `sourceFileName` (backward compat) |
| "No source workbook" message | `"Please re-upload..."` | `"No session found. Please upload first."` |
| Export check | `if (!sourceFileName)` | `if (!sessionId)` |

---

## Architecture After Phase 5.2

```
USER FLOW (same as ConMas):

Upload Excel
    │
    ▼
Server stores XLSX → TempWorkbooks/{sessionId}/original.xlsx
    │
    ▼
Returns { sessionId, previewUrl, runtimeForm }
    │
    ▼
Browser stores ONLY sessionId (no filenames)
    │
    ▼
Edit fields → Export Excel
    │
    ▼
POST /api/form/save-edited { sessionId, workbookDefinition }
    │
    ▼
Server resolves: SessionId → TempWorkbooks/{sessionId}/original.xlsx
    │
    ▼
WorkbookValueWriter edits ONLY cell values
    │
    ▼
WorkbookDiffValidator validates
    │
    ▼
Download edited XLSX
```

---

## Files Created/Modified

| File | Change |
|------|--------|
| `Application/SessionWorkbookStore.cs` | **NEW** — Server-side session storage |
| `Application/ServiceRegistration.cs` | Added `ISessionWorkbookStore` DI registration |
| `Application/FormSaveService.cs` | Added `SaveEditedValuesAsync(WbDef, dir, sourcePath)` overload |
| `Controllers/FormController.cs` | All upload endpoints return `sessionId`, save-edited resolves from session store |
| `Models/WorkbookDefinition/WorkbookDefinition.cs` | Added `SessionId`, deprecated `SourceFileName` |
| `paperless/app/page.tsx` | Replaced `sourceFileName` with `sessionId` throughout |

---

## Key Results

- ✅ **Browser never tracks filenames** — just a `sessionId`
- ✅ **Server owns the workbook** — same as ConMas Designer
- ✅ **No "No source workbook available" messages**
- ✅ **Backward compat preserved** — `SourceFileName` still works as fallback
- ✅ **Session expiry handled** — HTTP 410 with clear message
- ✅ **24-hour session lifetime** — files auto-cleaned on access
- ✅ **Build: 0 errors**

---

## Known Limitations

| # | Limitation | Impact |
|---|------------|--------|
| 1 | `GetRuntime()` still looks in `Forms/` — not updated for session store | OK: `FromExcel` and `UploadPreview` return runtime form directly in response |
| 2 | `CleanupExpiredSessions()` exists but not called by background timer | Abandoned sessions remain on disk until next `GetSession()` access |
| 3 | Session store is file-based (no database) | Sessions don't survive server redeploy (acceptable for dev/test) |

---

## Phase 6 Roadmap

| Phase | Focus |
|-------|-------|
| 6.0 | Field Insertion/Deletion |
| 6.1 | Conditional Formatting Editor |
| 6.2 | Multi-User Collaboration |
| 6.3 | Performance Optimization |
| 6.4 | Background session cleanup service (addresses limitation #2) |
