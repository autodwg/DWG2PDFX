# PHP Web Application - DWG to PDF Converter

A web-based GUI for DWG to PDF conversion using PHP.

## Prerequisites

1. Register the COM DLL (run as Administrator):
   - 64-bit: `regsvr32 "path\to\dwgtopdfx64.dll"`
   - 32-bit: `regsvr32 "path\to\dwgtopdfx.dll"`

2. PHP 7.0+ with COM extension enabled. Edit `php.ini`:
   ```
   extension=com_dotnet
   ```

## How to Run

```
cd "D:\AutoDWG SDK\DWG2PDFX\example\php-web"
php -S localhost:9877
```

Then open http://localhost:9877 in your browser.

## Features

- Upload DWG files via web browser
- Configure conversion parameters (color mode, DPI, page size, etc.)
- Collapsible settings panel
- Conversion result with popup notification
- Download converted PDF files

## Directory Structure

```
php-web/
├── index.php      # Main application (single file)
├── uploads/       # Uploaded DWG files (auto-created)
└── output/        # Converted PDF files (auto-created)
```

## Notes

- The PHP architecture (32/64-bit) must match the registered DLL.
- On Windows, use XAMPP, WAMP, or the built-in PHP server.