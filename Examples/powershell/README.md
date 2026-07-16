# PowerShell Example - DWG to PDF Conversion

## Prerequisites

1. Register the COM DLL (run as Administrator):
   - 64-bit: `regsvr32 "path\to\dwgtopdfx64.dll"`
   - 32-bit: `regsvr32 "path\to\dwgtopdfx.dll"`

2. Place your test DWG file in this directory or adjust the file paths in the script.

## How to Run

```powershell
powershell -ExecutionPolicy Bypass -File example.ps1
```

Or open PowerShell ISE and run the script directly.

## Examples Included

The file contains 4 functions:

1. `Convert-Basic` - Simple DWG to PDF conversion
2. `Convert-PerView` - One PDF per layout/view
3. `Convert-Batch` - Batch convert multiple files
4. `Convert-Advanced` - Full API demonstration

## Notes

- Uses `New-Object -ComObject DWGTOPDFX.ConvertPDF`.
- If you get a permission error, try running PowerShell as Administrator.