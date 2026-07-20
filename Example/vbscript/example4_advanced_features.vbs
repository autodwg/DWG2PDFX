' ============================================================
' AutoDWG DWG2PDF SDK - Example 4: Advanced Features
'
' Demonstrates:
'   - CTB color book mapping (SetCTBFile)
'   - Output scale ratio (SetOutputScale)
'   - ConvertPapers (without Model Space)
'   - Multiple pen widths
'   - Background color (nBackgroundColor)
'   - Exclude layers (Addexcudelayer)
'
' Prerequisites:
'   regsvr32 "path\to\dwgtopdfx64.dll"  (run as Administrator)
' ============================================================

Option Explicit

Dim obj
Set obj = CreateObject("DWGTOPDFX.ConvertPDF")

' --- Basic Parameters ---
obj.ConvertNotPlotLayer = 0
obj.Width  = 914                  ' Large format page width (mm)
obj.Height = 610                  ' Large format page height (mm)
obj.ColorMode = 1                 ' Black & White
obj.DPI = 100
obj.OutWithTTF = 1
obj.EmbeddedSearchText = 0
obj.ZoomType = 0
obj.nBackgroundColor = 7   ' White background

' --- Advanced: CTB Color Book ---
' Maps AutoCAD color indices to output line weights/colors
obj.SetCTBFile "acad.ctb"

' --- Advanced: Output Scale ---
obj.SetOutputScale "1"            ' 1:1 original scale

' --- Pen Widths for different color indices ---
obj.PenWidth(1) = 0.02            ' Red
obj.PenWidth(2) = 0.03            ' Yellow
obj.PenWidth(3) = 0.05            ' Green
obj.PenWidth(4) = 0.08            ' Cyan
obj.PenWidth(5) = 0.10            ' Blue
obj.PenWidth(6) = 0.15            ' Magenta
obj.PenWidth(7) = 0.20            ' White/Black

' --- Font & PS paths ---
obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"
obj.AddFontPath("C:\windows\system32\shxfont\")

' --- Input ---
obj.InputFile = "Office.dwg"

' --- ConvertPapers: converts layouts only (no Model Space) ---
WScript.Echo "Begin converting (papers only, no Model Space)..."
On Error Resume Next
obj.ConvertPapers("output_papers.pdf")
If Err.Number <> 0 Then
    WScript.Echo "Conversion failed. (Error: 0x" & Hex(Err.Number) & ")"
Else
    WScript.Echo "Conversion complete."
End If
On Error GoTo 0

Set obj = Nothing