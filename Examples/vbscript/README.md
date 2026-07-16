# VBScript Examples - DWG to PDF Conversion

## Prerequisites

1. Register the COM DLL (run as Administrator):
   - 64-bit: `regsvr32 "path\to\dwgtopdfx64.dll"`
   - 32-bit: `regsvr32 "path\to\dwgtopdfx.dll"`

2. Place your test DWG file in this directory or adjust the file paths in the scripts.

## Examples

| File | Description |
|------|-------------|
| `example1_basic_conversion.vbs` | Basic DWG to PDF conversion with default settings |
| `example2_one_pdf_per_view.vbs` | Export each view/layout to a separate PDF |
| `example3_batch_conversion.vbs` | Batch convert multiple DWG files |
| `example4_advanced_features.vbs` | Advanced options: font paths, CTB, output scale, layer control |

## How to Run

Double-click the `.vbs` file, or run from Command Prompt:

```
cscript example1_basic_conversion.vbs
```

## Notes

- VBScript uses `CreateObject("DWGTOPDFX.ConvertPDF")` to create the COM object.
- On 64-bit Windows, use the 64-bit DLL and run with `C:\Windows\System32\cscript.exe`.
- On 32-bit Windows, use the 32-bit DLL and run with `C:\Windows\SysWOW64\cscript.exe`.