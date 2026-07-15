# Build Phase 1 - Excel Print Area Capture API

## Overview

Phase 1 implements the foundational **Excel Print Area Capture** functionality. The API accepts Excel files (.xlsx/.xls), automates Microsoft Excel via COM Interop, reads the configured Print Area from the first worksheet, captures it as a PNG image, and returns the image URL to the frontend for display.

---

## Changes Made

### Backend — ASP.NET Core Web API (`ExcelAPI/`)

#### 1. NuGet Package Added
- **`Microsoft.Office.Interop.Excel`** (v15.0.4795.1001) — enables COM Interop with Microsoft Excel

#### 2. New Project Structure
```
ExcelAPI/ExcelAPI/
├── Controllers/
│   └── ExcelController.cs          # Upload endpoint
├── Services/
│   ├── Interfaces/
│   │   └── IExcelCaptureService.cs # Service contract
│   └── ExcelCaptureService.cs      # COM automation logic
├── Models/
│   ├── ApiResponse.cs              # Generic API response wrapper
│   └── UploadResponse.cs           # Image URL response model
├── Uploads/                        # Temporary uploaded files (auto-cleaned)
├── Preview/                        # Captured PNG images (static files)
├── Program.cs                      # DI, CORS, static files config
├── ExcelAPI.csproj                 # Updated with COM package
└── appsettings.json                # Unchanged
```

#### 3. Files Created

| File | Purpose |
|------|---------|
| `Controllers/ExcelController.cs` | POST `/api/excel/upload` — validates file, saves to Uploads, delegates to service, returns response |
| `Services/Interfaces/IExcelCaptureService.cs` | Interface with `Task<string> CapturePrintAreaAsync(string excelPath)` |
| `Services/ExcelCaptureService.cs` | Full Excel COM automation: launch Excel, open workbook, read PrintArea, capture range as PNG via Chart export, cleanup |
| `Models/ApiResponse.cs` | Generic `ApiResponse<T>` with Success, Message, and Data fields |
| `Models/UploadResponse.cs` | Response DTO with `ImageUrl` property |

#### 4. Files Modified

| File | Changes |
|------|---------|
| `Program.cs` | Added DI registration for `IExcelCaptureService`, CORS policy for `localhost:3000`, static file serving for `/preview` folder |
| `ExcelAPI.csproj` | Added `Microsoft.Office.Interop.Excel` package reference, suppressed CA1416 platform warnings |

#### 5. Files Deleted
- `Controllers/WeatherForecastController.cs` (sample template)
- `WeatherForecast.cs` (sample model)

#### 6. Upload Flow
1. **Validate** — rejects empty files and unsupported extensions (.xlsx/.xls only)
2. **Save** — stores file in `Uploads/` with a unique GUID-based filename
3. **Open Excel** — launches hidden Excel instance via COM Interop
4. **Open Workbook** — opens the uploaded workbook
5. **Read Print Area** — reads `worksheet.PageSetup.PrintArea` from first worksheet
6. **Validate Print Area** — returns error if no print area is configured
7. **Capture** — copies the Print Area range as a bitmap picture, creates a temporary chart sheet, pastes the picture, and exports as PNG
8. **Cleanup** — closes workbook, quits Excel, releases all COM objects via `Marshal.ReleaseComObject()`, forces garbage collection
9. **Return** — returns JSON with `imageUrl` (e.g., `/preview/page_abc123.png`)
10. **Clean uploaded file** — deletes the temp file from Uploads

#### 7. API Endpoint

```http
POST /api/excel/upload
Content-Type: multipart/form-data

Field: file (accepts .xlsx, .xls)
```

**Success Response (200):**
```json
{
  "success": true,
  "message": "Print area captured successfully.",
  "data": {
    "imageUrl": "/preview/page_abc123.png"
  }
}
```

**Error Response (400/500):**
```json
{
  "success": false,
  "message": "No print area is configured in this worksheet..."
}
```

#### 8. COM Cleanup
- Workbook is closed without saving changes
- Excel application is quit
- Every COM object is released via `Marshal.ReleaseComObject()`
- `GC.Collect()` and `GC.WaitForPendingFinalizers()` are called twice to ensure cleanup
- No orphaned EXCEL.EXE processes should remain

#### 9. Error Handling
| Scenario | HTTP Status | Message |
|----------|-------------|---------|
| No file uploaded | 400 | "No file uploaded..." |
| Invalid extension | 400 | "Unsupported file extension..." |
| No print area configured | 500 | "No print area is configured..." |
| Excel cannot start | 500 | "Failed to process the Excel file..." |
| COM exception | 500 | Wrapped with user-friendly message |

---

### Frontend — Next.js (`paperless/`)

#### Files Modified

| File | Changes |
|------|---------|
| `app/page.tsx` | Complete rewrite: file upload form, progress states, error/success messages, preview display with Image component |
| `app/layout.tsx` | Updated metadata title and description for PaperLess Enterprise |

#### Frontend Features
- File input restricted to `.xlsx` and `.xls` files
- Loading spinner during upload
- Success/error message display with icons
- Captured image preview with responsive layout
- "Upload Another" button to reset and upload again
- Header with app branding
- Footer with phase indicator

#### Environment Configuration
The API URL is configurable via `NEXT_PUBLIC_API_URL` environment variable (defaults to `http://localhost:5090`).

---

## Running the Application

### Backend
```bash
cd ExcelAPI/ExcelAPI
dotnet run
```
The API starts at `http://localhost:5090` by default.

### Frontend
```bash
cd paperless
npm run dev
```
The frontend starts at `http://localhost:3000`.

### Usage
1. Open `http://localhost:3000` in a browser
2. Select an Excel file (.xlsx/.xls) that has a configured Print Area
3. Click **Upload & Capture**
4. The captured Print Area will be displayed as a PNG image

---

## Out of Scope (Not Implemented)
- Custom rendering engine
- HTML/SVG/Canvas rendering
- PDF export
- OCR
- Database
- Authentication
- Multiple worksheets
- Multiple pages
- Form designer
- Field mapping
- Template management
