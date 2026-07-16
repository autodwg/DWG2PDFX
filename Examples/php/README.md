# PHP Example - DWG to PDF Conversion

## Prerequisites

1. Register the COM DLL (run as Administrator):
   - 64-bit: `regsvr32 "path\to\dwgtopdfx64.dll"`
   - 32-bit: `regsvr32 "path\to\dwgtopdfx.dll"`

2. PHP with COM extension enabled. Edit `php.ini` and uncomment:
   ```
   extension=com_dotnet
   ```

3. Place your test DWG file in this directory or adjust the file paths.

## How to Run

```
php example.php
```

## Examples Included

The file contains 4 functions:

1. `basicConversion()` - Simple DWG to PDF
2. `onePdfPerView()` - Export each view/layout
3. `batchConversion()` - Convert multiple files
4. `advancedFeatures()` - Full API demonstration

## Notes

- Uses `new COM("DWGTOPDFX.ConvertPDF")`.
- The PHP architecture (32/64-bit) must match the registered DLL.
- If you get "COM object not found", check that `extension=com_dotnet` is enabled in `php.ini`.