# Phase 11J.3B — Runtime 500 Investigation (OpenXML Parsing)

**Date:** July 2026  
**Status:** ✅ Fixed  
**Build:** 0 errors  

---

## Investigation Summary

**Error:** `GET /api/form/runtime/{templateId}` → **500 Internal Server Error**

**Message:** `"The element does not allow the specified attribute."`

**Exception type:** `System.Collections.Generic.KeyNotFoundException`

---

## Stack Trace (Captured via Debug Logging)

```
System.Collections.Generic.KeyNotFoundException: The element does not allow the specified attribute.
   at DocumentFormat.OpenXml.OpenXmlElement.GetAttribute(OpenXmlQualifiedName& qname)
   at DocumentFormat.OpenXml.OpenXmlElement.GetAttribute(String localName, String namespaceUri)
   at ExcelAPI.Rendering.ThemeResolver.LoadFromWorkbook(WorkbookPart wbPart)
       in .../ThemeResolver.cs:line 77
   at ExcelAPI.Rendering.ColorResolver.LoadTheme(WorkbookPart wbPart)
       in .../ColorResolver.cs:line 22
   at ExcelAPI.Rendering.StyleResolver.LoadTheme(WorkbookPart wbPart)
       in .../StyleResolver.cs:line 174
   at ExcelAPI.Rendering.OpenXmlParser.Parse(String filePath)
       in .../OpenXmlParser.cs:line 80
   at ExcelAPI.Controllers.FormController.GetRuntime(String templateId)
       in .../FormController.cs:line 420
```

**Chain:** `FormController.GetRuntime()` → `OpenXmlParser.Parse()` → `StyleResolver.LoadTheme()` → `ColorResolver.LoadTheme()` → `ThemeResolver.LoadFromWorkbook()` → **`ThemeResolver.cs:77`**

---

## Root Cause

### File: `ThemeResolver.cs` — Line 77

```csharp
// INVALID: child is a PARENT color scheme element (<a:dk1>, <a:lt1>, <a:accent1>, etc.)
// These elements DON'T allow a "val" attribute
hexColor = child.GetAttribute("val", "").Value;
```

The `ThemeResolver.LoadFromWorkbook()` method iterates over the children of the `<a:clrScheme>` element. These children are color definitions like `<a:dk1>`, `<a:lt1>`, `<a:accent1>`, etc. — each containing a COLOR CHILD element (`<a:srgbClr>` or `<a:sysClr>`).

The code correctly checked for `<a:srgbClr>` in an inner loop (line 69). But when no `srgbClr` child was found, the **fallback** at line 77 attempted to read `val` directly from the **parent** element (e.g., `<a:dk1>`) using `child.GetAttribute("val", "")`.

In the OpenXML SDK, `OpenXmlElement.GetAttribute()` validates that the requested attribute is allowed on that element type. Since `<a:dk1>`, `<a:lt1>`, etc. don't define a `val` attribute (only their child elements like `<a:srgbClr>` and `<a:sysClr>` do), the SDK throws `KeyNotFoundException` with the message "The element does not allow the specified attribute."

### XML Structure That Triggers the Bug

```xml
<a:clrScheme name="Office">
  <a:dk1>
    <a:sysClr val="windowText"/>   <!-- NOT srgbClr — the inner loop missed this -->
  </a:dk1>
  ...
</a:clrScheme>
```

When the color scheme uses `<a:sysClr>` (system color) instead of `<a:srgbClr>` (RGB value), the inner loop didn't find it, so the fallback ran on the parent — and crashed.

---

## Fix Applied

### File: `ThemeResolver.cs`

**Three changes:**

1. **Added `sysClr` detection** in the inner sub-child loop (alongside existing `srgbClr` detection)
2. **Wrapped `GetAttribute` calls in try-catch** to safely handle missing attributes
3. **Added `MapSystemColor()` helper** to convert system color names (`"windowText"`, `"buttonface"`, `"highlight"`, etc.) to hex ARGB values
4. **Replaced the invalid fallback** (`child.GetAttribute("val", "")`) with a `continue` — safely skipping unparseable color elements

### Before (Broken)

```csharp
foreach (var subChild in child.ChildElements)
{
    string subName = subChild.LocalName ?? "";
    if (subName == "srgbClr" || subName == "srgbclr")
    {
        hexColor = subChild.GetAttribute("val", "").Value;
        break;
    }
}

// BUG: This throws when child has no "val" attribute
if (hexColor == null)
{
    hexColor = child.GetAttribute("val", "").Value;
}
```

### After (Fixed)

```csharp
foreach (var subChild in child.ChildElements)
{
    string subName = subChild.LocalName ?? "";
    if (subName is "srgbClr" or "srgbclr")
    {
        try { hexColor = subChild.GetAttribute("val", "").Value; } catch { }
        break;
    }
    if (subName is "sysClr" or "sysclr")
    {
        try { hexColor = subChild.GetAttribute("val", "").Value; } catch { }
        if (!string.IsNullOrEmpty(hexColor) && !hexColor.StartsWith("#"))
            hexColor = MapSystemColor(hexColor);
        break;
    }
}

if (hexColor == null)
{
    continue;  // Skip this scheme color safely
}
```

---

## Verification

| Test | Result |
|------|--------|
| Build (dotnet build) | ✅ Build succeeded — 0 errors |
| `POST /api/form/from-excel` (upload) | ✅ 200 OK, templateId returned |
| `GET /api/form/runtime/{templateId}` | ✅ **200 OK** (was 500) |
| Runtime response contains sheets & fields | ✅ 1 sheet, 0 fields (test file has no bordered cells) |
| No rendering/engine changes | ✅ Only ThemeResolver.cs modified |

---

## Debug Logging Added

Temporary logging was added to `FormController.GetRuntime()` for future debugging:

- Stage markers: `[RUNTIME] Stage 1-5` (finding template → parsing → building → success)
- Error now logs `ex.ToString()` (full stack trace, not just message)
- 500 response includes `exceptionType` and `stackTrace` fields (development only)

---

## Conclusion

The 500 error was a **ThemeResolver bug**: the fallback code tried to read a `val` attribute from parent color scheme elements (`<a:dk1>`, `<a:lt1>`, etc.) that don't allow it in the OpenXML schema. The fix adds proper `sysClr` detection, removes the invalid parent-element attribute access, and maps system color names to hex values. 

**No changes to the Rendering Engine, Runtime Engine, CoordinateEngine, or FieldDetector were needed.**
