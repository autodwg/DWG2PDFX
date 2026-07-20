' ============================================================
' AutoDWG DWG2PDF SDK - Example 2: One PDF Per Layout/View
'
' Iterates through all views/layouts in a DWG file and
' converts each one to a separate PDF file.
' Demonstrates SetOutColor for color index mapping and
' Addexcudelayer for excluding layers.
'
' Prerequisites:
'   regsvr32 "path\to\dwgtopdfx64.dll"  (run as Administrator)
' ============================================================

Option Explicit

Dim obj
Set obj = CreateObject("DWGTOPDFX.ConvertPDF")

' --- Configuration ---
obj.Addexcudelayer "Defpoints"    ' Exclude the Defpoints layer
obj.ConvertNotPlotLayer = 0       ' Skip non-printing layers

obj.Width = 300                   ' Page width in mm
obj.Height = 200                  ' Page height in mm
obj.ColorMode = 1                 ' Black & White mode

' Set pen widths for different color indices
obj.PenWidth(1) = 0.04
obj.PenWidth(2) = 0.04
obj.PenWidth(3) = 0.04
obj.PenWidth(4) = 0.04
obj.PenWidth(5) = 0.04
obj.PenWidth(6) = 0.04
obj.PenWidth(7) = 0.04

' --- Color mapping: map input color index to output color ---
' SetOutColor(inputColorIndex, outputColorIndex)
' Example: map all red (1) lines to blue (5) in output
obj.SetOutColor 1, 5
' obj.SetOutColor 2, 3            ' Map yellow to green
' obj.SetOutColor 7, 1            ' Map white/black to red

obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"
obj.InputFile = "Office.dwg"

' --- Iterate through all views ---
Dim I, strViewName, strOutput
WScript.Echo "Total " & obj.ViewCount & " view(s) found."

On Error Resume Next
For I = 1 To obj.ViewCount
    strViewName = obj.ViewName(I)
    
    ' Display page setup info
    WScript.Echo "  [" & I & "] View: " & strViewName & _
           ", Setup: " & obj.PageSetupName(I) & _
           ", Unit: " & obj.PageSizeUnit(I) & _
           ", Size: " & obj.PageWidth(I) & "x" & obj.PageHeight(I) & "mm"
    
    ' Convert current view to PDF
    strOutput = strViewName & ".pdf"
    obj.ConvertView strOutput, strViewName
    
    If Err.Number <> 0 Then
        WScript.Echo "  Failed to convert view: " & strViewName & " (Error: 0x" & Hex(Err.Number) & ")"
        Err.Clear
    End If
Next
On Error GoTo 0

WScript.Echo "Done."

Set obj = Nothing