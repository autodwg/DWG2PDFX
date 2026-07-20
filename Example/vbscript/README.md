# VBScript Examples - DWG to PDF Conversion

## Prerequisites

1. Register the COM DLL (run as Administrator):
   - 64-bit: `regsvr32 "path\to\dwgtopdfx64.dll"`
   - 32-bit: `regsvr32 "path\to\dwgtopdfx.dll"`

2. Place your test DWG file in this directory or adjust the file paths in the scripts.

## Examples

| File | Description |
|------|-------------|
| `example1_basic_conversion.vbs` | Basic DWG to PDF with all settings (background color, DPI) |
| `example2_one_pdf_per_view.vbs` | Export each view/layout to a separate PDF, with SetOutColor |
| `example3_batch_conversion.vbs` | Batch convert folder (per-view + AddFile/ConvertFiles) |
| `example4_advanced_features.vbs` | CTB, output scale, ConvertPapers, background color |
| `example5_dwf_to_dwg.vbs` | Convert DWF file to DWG/DXF format |

## How to Run

Double-click the `.vbs` file, or run from Command Prompt:

```
cscript example1_basic_conversion.vbs
```

## Notes

- VBScript uses `CreateObject("DWGTOPDFX.ConvertPDF")` to create the COM object.
- On 64-bit Windows, use the 64-bit DLL and run with `C:\Windows\System32\cscript.exe`.
- On 32-bit Windows, use the 32-bit DLL and run with `C:\Windows\SysWOW64\cscript.exe`.

## API Reference

All examples call the same COM interface with ProgID `DWGTOPDFX.ConvertPDF`.

### Properties

| Property | Type | Access | Description |
|----------|------|--------|-------------|
| `ColorMode` | int | write | Color mode: 0=256 colors, 1=B&W |
| `Width` | double | write | Page width in mm |
| `Height` | double | write | Page height in mm |
| `DPI` | int | write | Output resolution |
| `DPIWidth` | short | write | Separate horizontal DPI |
| `DPIHeight` | short | write | Separate vertical DPI |
| `InputFile` | string | write | Input DWG file path |
| `ViewCount` | int | read | Number of views/layouts |
| `ViewName` | string | read | View name (indexed) |
| `PenWidth` | double | read/write | Pen width (indexed by color) |
| `PageWidth` | double | read | Page width (indexed) |
| `PageHeight` | double | read | Page height (indexed) |
| `PageSizeUnit` | short | read | Page unit (indexed) |
| `PageSetupName` | string | read | Page setup name (indexed) |
| `PSPath` | string | write | PostScript search path |
| `OutWithTTF` | short | read/write | Output TrueType fonts |
| `EmbeddedSearchText` | short | read/write | Embed searchable text |
| `ZoomType` | short | read/write | Zoom: 0=Extents, 1=Last view |
| `OutputByLayoutSize` | bool | read/write | Output by layout dimensions |
| `nBackgroundColor` | short | read/write | Background color (RGB integer) |
| `ConvertNotPlotLayer` | int | read/write | Convert non-plottable layers |

### Methods

| Method | Parameters | Description |
|--------|-----------|-------------|
| `Convert` | `(string szoutFile)` | Convert all layouts to PDF |
| `ConvertView` | `(string szOutputFileName, ref object view)` | Convert a specific view |
| `ConvertPapers` | `(string szoutFile)` | Convert paper space only (no Model) |
| `AddFontPath` | `(string szFontPath)` | Add font search path |
| `AddExcudeLayer` | `(string szLayer)` | Exclude a layer from output |
| `SetOutColor` | `(int nColorIndex, int nOutColorIndex)` | Map input color to output color |
| `SetCTBFile` | `(string szCTBFile)` | Set CTB plot style file |
| `SetOutputScale` | `(float dScale)` | Set output scale factor |
| `AddFile` | `(string szFile)` | Add file for batch conversion |
| `ConvertFiles` | `(string szOutputFile)` | Batch convert added files |
| `ConvertDWFtoDWGDXF` | `(string strDWFIn, string strOutputFile, short nOutVersion)` | Convert DWF to DWG/DXF |