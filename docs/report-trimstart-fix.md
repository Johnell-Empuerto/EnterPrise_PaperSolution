# TrimStart Fix Report

## Problem
In .NET 10, `string.TrimStart(string)` returns `ReadOnlySpan<char>` instead of `string`, causing CS0029 compilation errors in `FormController.cs` where the result was passed to `Path.GetFileName(string)`.

## Affected File
`ExcelAPI/ExcelAPI/Controllers/FormController.cs`

## Changes Made
Two occurrences of the following pattern were replaced:

### Before (both occurrences)
```csharp
string previewFileName = Path.GetFileName(
    capture.ImageUrl.TrimStart('/').TrimStart("preview/"));
```

### After - Occurrence 1 (thumbnail generation, ~line 207)
```csharp
string rawUrl = capture.ImageUrl.TrimStart('/');
if (rawUrl.StartsWith("preview/"))
    rawUrl = rawUrl.Substring("preview/".Length);
string previewFileName = Path.GetFileName(rawUrl);
```

### After - Occurrence 2 (background persistence, ~line 256)
```csharp
string rawUrl = capture.ImageUrl.TrimStart('/');
if (rawUrl.StartsWith("preview/"))
    rawUrl = rawUrl.Substring("preview/".Length);
string previewFileName2 = Path.GetFileName(rawUrl);
```

## Build Result
- **0 errors**, 2 NuGet vulnerability warnings (pre-existing, unrelated)
- Build: **Succeeded**

## Root Cause
A .NET 10 behavioral change where `TrimStart(ReadOnlySpan<char>)` overload replaces the `TrimStart(string)` overload, returning `ReadOnlySpan<char>` instead of `string`.

## Resolution
Replaced `TrimStart("preview/")` with an equivalent `StartsWith`/`Substring` pattern that works across all .NET versions.
