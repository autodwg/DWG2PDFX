' ============================================================
' AutoDWG DWG2PDF SDK - Example 5: DWF to DWG/DXF Conversion
'
' Demonstrates the ConvertDWFtoDWGDXF method which converts
' DWF files to DWG or DXF format.
'
' Parameters:
'   strDWFIn     - Input DWF file path
'   strOutputFile - Output DWG or DXF file path
'   nOutVersion  - Output version (e.g., AutoCAD version code)
'
' Prerequisites:
'   regsvr32 "path\to\dwgtopdfx64.dll"  (run as Administrator)
' ============================================================

Option Explicit

Dim obj
Set obj = CreateObject("DWGTOPDFX.ConvertPDF")

Dim inputFile, outputFile, outVersion

inputFile = "drawing.dwf"
outputFile = "output.dwg"
outVersion = 24                 ' Output DWG version code

WScript.Echo "Converting DWF to DWG..."
WScript.Echo "  Input:  " & inputFile
WScript.Echo "  Output: " & outputFile

On Error Resume Next
obj.ConvertDWFtoDWGDXF inputFile, outputFile, outVersion
If Err.Number <> 0 Then
    WScript.Echo "Conversion failed. (Error: 0x" & Hex(Err.Number) & ")"
Else
    WScript.Echo "Conversion complete."
End If
On Error GoTo 0

' --- Convert to DXF instead ---
' outputFile = "output.dxf"
' obj.ConvertDWFtoDWGDXF inputFile, outputFile, outVersion

Set obj = Nothing