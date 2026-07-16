# Python Example - DWG to PDF Conversion

## Prerequisites

1. Register the COM DLL (run as Administrator):
   - 64-bit: `regsvr32 "path\to\dwgtopdfx64.dll"`
   - 32-bit: `regsvr32 "path\to\dwgtopdfx.dll"`

2. Install the `pywin32` package:
   ```
   pip install pywin32
   ```

3. Place your test DWG file in this directory or adjust the file paths in the script.

## How to Run

```
python example.py
```

## Examples Included

The file contains 4 functions:

1. `example_basic_conversion()` - Simple DWG to PDF
2. `example_one_pdf_per_view()` - Export each view/layout separately
3. `example_batch_conversion()` - Convert multiple files
4. `example_advanced_features()` - Full API usage

## Notes

- Uses `win32com.client.Dispatch("DWGTOPDFX.ConvertPDF")`.
- Make sure the Python architecture (32/64-bit) matches the registered DLL.