# C# WinForms Application - DWG to PDF Converter

A desktop GUI application for DWG to PDF conversion using Windows Forms.

## Prerequisites

1. Register the COM DLL (run as Administrator):
   - 64-bit: `regsvr32 "path\to\dwgtopdfx64.dll"`
   - 32-bit: `regsvr32 "path\to\dwgtopdfx.dll"`

2. .NET 8.0 SDK

## How to Build and Run

```
cd "D:\www.autodwg.com\Server64\DWG2PDFConverter"
dotnet run
```

## Features

- Browse and select DWG input files
- Choose PDF output path
- Configure conversion parameters (color mode, DPI, page size, etc.)
- Real-time conversion log
- Popup notification on completion with option to open output folder

## Notes

- This is a Windows-only application (WinForms).
- The project file is located in the Server64 workspace at `DWG2PDFConverter\`.