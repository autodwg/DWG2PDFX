' ============================================================
' AutoDWG DWG2PDF SDK - Example 2: One PDF Per Layout/View
'
' Iterates through all views/layouts in a DWG file and
' converts each one to a separate PDF file.
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

obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"
obj.InputFile = "Office.dwg"

' --- Iterate through all views ---
Dim I, strViewName, strOutput
MsgBox "Total " & obj.ViewCount & " view(s) found."

For I = 1 To obj.ViewCount
    strViewName = obj.ViewName(I)
    
    ' Display page setup info for debugging
    MsgBox "Converting view: " & strViewName & vbCrLf & _
           "Page Setup: " & obj.PageSetupName(I) & vbCrLf & _
           "Unit: " & obj.PageSizeUnit(I) & vbCrLf & _
           "Width: " & obj.PageWidth(I) & " Height: " & obj.PageHeight(I)
    
    ' Convert current view to PDF
    strOutput = strViewName & ".pdf"
    obj.ConvertView strOutput, strViewName
Next

If Err.Number < 0 Then
    MsgBox "Error: " & Err.Description
End If

Set obj = Nothing