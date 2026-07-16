# C# Example - DWG to PDF Conversion

## Prerequisites

1. Register the COM DLL (run as Administrator):
   - 64-bit: `regsvr32 "path\to\dwgtopdfx64.dll"`
   - 32-bit: `regsvr32 "path\to\dwgtopdfx.dll"`

2. .NET Framework 4.0+ or .NET Core / .NET 5+

## How to Build and Run

### Using csc (C# Compiler)

```
csc /out:Example.exe Example.cs
Example.exe
```

### Using Visual Studio

1. Create a new Console Application project.
2. Add `Example.cs` to the project.
3. Build and run.

## Examples Included

The file contains 4 scenarios (uncomment the one you want to run):

1. **Basic Conversion** - Simple DWG to PDF with default settings
2. **One PDF per View** - Export each layout/view to a separate PDF
3. **Batch Conversion** - Convert multiple DWG files in a loop
4. **Advanced Features** - Font paths, CTB files, output scale, layer exclusion

## Notes

- Uses `Type.GetTypeFromProgID("DWGTOPDFX.ConvertPDF")` for late binding.
- The `dynamic` keyword requires C# 4.0+.