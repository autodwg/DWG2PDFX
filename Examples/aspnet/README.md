# ASP.NET Core Web Application - DWG to PDF Converter

A web-based GUI for DWG to PDF conversion using ASP.NET Core Razor Pages.

## Prerequisites

1. Register the COM DLL (run as Administrator):
   - 64-bit: `regsvr32 "path\to\dwgtopdfx64.dll"`
   - 32-bit: `regsvr32 "path\to\dwgtopdfx.dll"`

2. .NET 8.0 SDK

## How to Build and Run

```
cd "D:\AutoDWG SDK\DWG2PDFX\example\aspnet"
dotnet run
```

The application will start on http://localhost:9876 and automatically open your browser.

## Features

- Upload DWG files via web browser
- Configure conversion parameters (color mode, DPI, page size, etc.)
- Collapsible settings panel
- Conversion result with popup notification
- Download converted PDF files via green download button

## Directory Structure

```
aspnet/
├── Program.cs             # Application entry point + COM service
├── DWG2PDFWeb.csproj      # Project file
├── Pages/
│   ├── Index.cshtml       # Main page (upload form + result)
│   ├── Index.cshtml.cs    # Page handler (upload + convert + download)
│   ├── _Layout.cshtml     # HTML layout + CSS styles
│   ├── _ViewImports.cshtml
│   └── _ViewStart.cshtml
├── uploads/               # Uploaded DWG files (auto-created)
└── output/                # Converted PDF files (auto-created)
```

## Notes

- This is a Windows-only application (requires COM interop).
- The target framework is `net8.0-windows`.
- Default port is 9876. Change in `Program.cs` if needed.