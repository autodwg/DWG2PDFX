# Python GUI Application - DWG to PDF Converter

A desktop GUI application for DWG to PDF conversion using Python tkinter.

## Prerequisites

1. Register the COM DLL (run as Administrator):
   - 64-bit: `regsvr32 "path\to\dwgtopdfx64.dll"`
   - 32-bit: `regsvr32 "path\to\dwgtopdfx.dll"`

2. Python 3.7+ with `pywin32`:
   ```
   pip install pywin32
   ```

## How to Run

```
python example_gui.py
```

## Features

- Browse and select DWG input files
- Choose PDF output path
- Configure conversion parameters (color mode, DPI, page size, etc.)
- Real-time conversion log
- Popup notification on completion

## Notes

- Uses Python's built-in `tkinter` for the GUI (no extra install needed).
- Windows only (requires COM support via `pywin32`).