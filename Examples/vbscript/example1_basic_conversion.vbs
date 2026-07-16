' ============================================================
' AutoDWG DWG2PDF SDK - Example 1: Basic Single File Conversion
'
' Converts all layouts of a DWG file into a single PDF.
'
' Prerequisites:
'   regsvr32 "path\to\dwgtopdfx64.dll"  (run as Administrator)
' ============================================================

Option Explicit

Dim obj
Set obj = CreateObject("DWGTOPDFX.ConvertPDF")

' --- Conversion Parameters ---
obj.ConvertNotPlotLayer = 0       ' 0=Skip non-printing layers (default), 1=Convert them
obj.Width  = 240                  ' Page width in mm
obj.Height = 180                  ' Page height in mm
obj.ColorMode = 0                 ' 0=256 colors, 1=Black & White
obj.DPI = 300                     ' Output resolution
obj.OutWithTTF = 1                ' 1=Preserve TrueType fonts as text
obj.EmbeddedSearchText = 0        ' 0=Keep text searchable, 1=Embed as image
obj.zoomtype = 0                  ' 0=Zoom extents, 1=Last view
obj.OutputByLayoutSize = 0        ' 0=Use Width/Height, 1=Use layout dimensions

' Pen width for color index 1 (Red), unit: mm
obj.PenWidth(1) = 0.02

' Font and PostScript support paths
obj.AddFontPath("C:\windows\system32\shxfont")
obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"

' --- Input/Output ---
obj.InputFile = "Office.dwg"

MsgBox "Begin converting..."
obj.Convert("output.pdf")
MsgBox "Conversion complete"

If Err.Number < 0 Then
    MsgBox "Error: " & Err.Description
End If

Set obj = Nothing