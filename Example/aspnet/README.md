# ASP.NET Core Web - DWG to PDF Converter

## Prerequisites

1. Register the COM DLL: `regsvr32 "path\to\dwgtopdfx64.dll"`
2. .NET 8.0 SDK

## Features

- Upload DWG files via browser
- Collapsible settings panel
- Background color configuration
- Auto-collapse after conversion, scroll to result
- Download PDF via link
- Popup notification

## How to Run

```
dotnet run
```

Access at: `http://localhost:9876`

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