# Node.js Example - DWG to PDF Conversion

## Prerequisites

1. Register the COM DLL (run as Administrator):
   - 64-bit: `regsvr32 "path\to\dwgtopdfx64.dll"`
   - 32-bit: `regsvr32 "path\to\dwgtopdfx.dll"`

2. Node.js 14+
3. Install dependencies:
   ```
   npm install
   ```

4. Place your test DWG file in this directory or adjust the file paths.

## How to Run

```
node example.js
```

## Examples Included

The file contains 4 functions:

1. `basicConversion()` - Simple DWG to PDF
2. `onePdfPerView()` - Export each view/layout
3. `batchConversion()` - Convert multiple files
4. `advancedFeatures()` - Full API usage

## Notes

- Uses the `win32-ole` npm package for COM interop.
- The Node.js architecture (32/64-bit) must match the registered DLL.