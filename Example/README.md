# DWG2PDFX SDK Examples

## Overview

This directory contains complete usage examples for the **DWG2PDFX** COM component across 12 programming languages and frameworks. Each example demonstrates how to convert DWG files to PDF, as well as DWF files to DWG/DXF format.

## Prerequisites

1. **Register the COM DLL** (run as Administrator):
   - 64-bit: `regsvr32 "path\to\dwgtopdfx64.dll"`
   - 32-bit: `regsvr32 "path\to\dwgtopdfx.dll"`

2. See each subdirectory's README.md for language-specific runtime requirements.

---

## Directory Structure

| Directory | Language/Framework | Description |
|-----------|-------------------|-------------|
| `vbscript/` | VBScript | Basic COM invocation, no extra dependencies |
| `csharp/` | C# (.NET) | Late binding via `dynamic` keyword |
| `python/` | Python | COM automation via `win32com.client` |
| `powershell/` | PowerShell | Windows scripting via `New-Object -ComObject` |
| `java/` | Java | COM bridge via JACOB library |
| `php/` | PHP (CLI) | PHP COM extension |
| `nodejs/` | Node.js | COM via `node-win32ole` |
| `go/` | Go | COM via `go-ole` library |
| `aspnet/` | ASP.NET Core | Web app with browser-based upload/convert UI |
| `csharp-winform/` | C# WinForms | Desktop GUI application |
| `python-gui/` | Python + tkinter | Desktop GUI application |
| `php-web/` | PHP built-in server | Web app with upload/convert UI |

---

## Five Core Examples

Each script/code file contains 5 examples demonstrating progressive complexity:

### Example 1: Basic Single File Conversion

**What it does**: Converts all layouts of a DWG file into a single PDF.

**APIs demonstrated**:
- `InputFile` — Input DWG file path
- `Width` / `Height` — Output page size (mm)
- `ColorMode` — Color mode (0=256 colors, 1=Black & White)
- `DPI` — Output resolution
- `DPIWidth` / `DPIHeight` — Independent horizontal/vertical DPI (optional)
- `OutWithTTF` — Preserve TrueType fonts as text
- `EmbeddedSearchText` — Embed text as image (not searchable)
- `ZoomType` — Zoom method (0=Zoom Extents, 1=Last View)
- `OutputByLayoutSize` — Use actual layout dimensions
- `nBackgroundColor` — Background color (ACI color index; 7=white, 0=black)
- `PenWidth(index)` — Line weight for a color index (mm)
- `AddFontPath()` — Add SHX font search path
- `PSPath` — PostScript resource path
- `ConvertNotPlotLayer` — Include non-plot layers
- `Convert()` — Execute conversion

**Use case**: Quick single DWG-to-PDF conversion, the most common scenario.

---

### Example 2: One PDF Per Layout/View

**What it does**: Iterates through all views/layouts in a DWG file and exports each one as a separate PDF.

**APIs demonstrated**:
- `ViewCount` — Total number of views/layouts (read-only)
- `ViewName(index)` — Name of a specific view (read-only)
- `PageSetupName(index)` — Page setup name (read-only)
- `PageWidth(index)` / `PageHeight(index)` — Page dimensions (read-only)
- `PageSizeUnit(index)` — Page size unit (read-only)
- `ConvertView(outputFile, viewName)` — Convert a specific view to PDF
- `SetOutColor(inputColorIndex, outputColorIndex)` — Remap color indices in output
- `Addexcudelayer(layerName)` — Exclude a layer from output

**Use case**: DWG files with multiple layouts (floor plan, elevation, section) that need separate output files.

---

### Example 3: Batch Folder Conversion

**What it does**: Converts all DWG files in a folder to PDF in batch.

**Two approaches demonstrated**:

**Approach A — File-by-file**:
- Iterate the folder, set `InputFile` for each DWG, call `ConvertView` per view
- Output naming: `DWGName_ViewName.pdf`

**Approach B — Batch queue**:
- `AddFile(filePath)` — Add DWG files to the batch queue
- `ConvertFiles(outputPath)` — Convert all queued files at once

**Use case**: Server-side batch processing of large volumes of DWG drawings.

---

### Example 4: Advanced Features

**What it does**: Demonstrates CTB color mapping, output scale, and papers-only conversion.

**APIs demonstrated**:
- `SetCTBFile(ctbFileName)` — Set CTB plot style table for color/lineweight mapping
- `SetOutputScale(scale)` — Set output scale ratio (e.g., 1.0 for 1:1)
- `ConvertPapers(outputFile)` — Convert only paper space layouts (excludes Model Space)

**Use case**: Precise scale control, CTB plot styles, or paper-space-only output.

---

### Example 5: DWF to DWG/DXF Conversion

**What it does**: Converts DWF files to DWG or DXF format.

**APIs demonstrated**:
- `ConvertDWFtoDWGDXF(inputFile, outputFile, outVersion)`
  - `inputFile` — Input DWF file path
  - `outputFile` — Output path (extension `.dwg` or `.dxf` determines format)
  - `outVersion` — Output DWG version code (e.g., 24 = AutoCAD 2020)

**Use case**: Restoring DWF reference files back to editable DWG/DXF format.

---

## Web Application Examples

### ASP.NET Core Web App (aspnet/)

- ASP.NET Core 8.0 Razor Pages application
- Browser UI: upload DWG, configure parameters, convert, download PDF
- Auto-opens browser at `http://localhost:9876` on startup
- Run: `dotnet run`

### PHP Web App (php-web/)

- PHP built-in development server
- Same functionality as the ASP.NET version
- Access at: `http://localhost:9877`
- Run: `php -S localhost:9877`

---

## Desktop GUI Examples

### C# WinForms (csharp-winform/)

- .NET 8.0 WinForms desktop application
- Full parameter configuration UI with logging output
- File dialogs, open output folder on completion
- Run: `dotnet run`

### Python + tkinter (python-gui/)

- Python tkinter desktop GUI
- Same functionality as the WinForms version
- Dependency: `pywin32`
- Run: `python example_gui.py`

---

## Complete API Quick Reference

### Properties

| Property | Type | R/W | Description |
|----------|------|-----|-------------|
| `InputFile` | string | Write | Input DWG file path |
| `ColorMode` | int | Write | 0=256 colors, 1=Black & White |
| `Width` | int | Write | Page width (mm) |
| `Height` | int | Write | Page height (mm) |
| `DPI` | int | Write | Uniform output resolution |
| `DPIWidth` | int | Write | Horizontal DPI (overrides DPI) |
| `DPIHeight` | int | Write | Vertical DPI (overrides DPI) |
| `OutWithTTF` | int | Write | 1=Preserve TrueType fonts as text |
| `EmbeddedSearchText` | int | Write | 1=Embed text as image |
| `ZoomType` | int | Write | 0=Zoom Extents, 1=Last View |
| `OutputByLayoutSize` | int | Write | 1=Use actual layout dimensions |
| `nBackgroundColor` | int | Write | Background color (ACI color index\) |
| `ConvertNotPlotLayer` | int | Write | 1=Convert non-plot layers |
| `PSPath` | string | Write | PostScript resource path |
| `ViewCount` | int | Read-only | Total number of views/layouts |
| `ViewName(index)` | string | Read-only | View name by index |
| `PenWidth(index)` | double | R/W | Line weight per color index (mm) |
| `PageWidth(index)` | double | Read-only | Page width |
| `PageHeight(index)` | double | Read-only | Page height |
| `PageSizeUnit(index)` | int | Read-only | Page size unit |
| `PageSetupName(index)` | string | Read-only | Page setup name |

### Methods

| Method | Description |
|--------|-------------|
| `Convert(outputFile)` | Convert all layouts to a single PDF |
| `ConvertView(outputFile, viewName)` | Convert a specific view to PDF |
| `ConvertPapers(outputFile)` | Convert paper space only (no Model Space) |
| `ConvertFiles(outputPath)` | Batch convert all queued files |
| `AddFile(filePath)` | Add a file to the batch queue |
| `AddFontPath(path)` | Add SHX font search path |
| `Addexcudelayer(layerName)` | Exclude a layer from output |
| `SetOutColor(inputIndex, outputIndex)` | Remap color indices |
| `SetCTBFile(fileName)` | Set CTB plot style table |
| `SetOutputScale(scale)` | Set output scale (float) |
| `ConvertDWFtoDWGDXF(input, output, version)` | Convert DWF to DWG/DXF |