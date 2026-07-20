' ============================================================
' AutoDWG DWG2PDF SDK - Example 1: Basic Single File Conversion
'
' Converts all layouts of a DWG file into a single PDF.
' Demonstrates all basic properties including background color
' and separate DPI width/height.
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
obj.DPI = 300                     ' Output resolution (uniform)
' obj.DPIWidth = 300              ' (Optional) Separate horizontal DPI
' obj.DPIHeight = 300             ' (Optional) Separate vertical DPI
obj.OutWithTTF = 1                ' 1=Preserve TrueType fonts as text
obj.EmbeddedSearchText = 0        ' 0=Keep text searchable, 1=Embed as image
obj.ZoomType = 0                  ' 0=Zoom extents, 1=Last view
obj.OutputByLayoutSize = 0        ' 0=Use Width/Height, 1=Use layout dimensions
obj.nBackgroundColor = 7   ' Background color (ACI: 7=white, 0=black)

' Pen width for color index 1 (Red), unit: mm
obj.PenWidth(1) = 0.02

' Font and PostScript support paths
obj.AddFontPath("C:\windows\system32\shxfont")
obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"

' --- Input/Output ---
obj.InputFile = "Office.dwg"

WScript.Echo "Begin converting..."
On Error Resume Next
obj.Convert("output.pdf")
If Err.Number <> 0 Then
    WScript.Echo "Conversion failed. (Error: 0x" & Hex(Err.Number) & ")"
Else
    WScript.Echo "Conversion complete."
End If
On Error GoTo 0

Set obj = Nothing