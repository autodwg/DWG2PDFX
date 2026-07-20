' ============================================================
' AutoDWG DWG2PDF SDK - Example 3: Batch Folder Conversion
'
' Two approaches demonstrated:
'   A) Iterate folder, convert each DWG with all views
'   B) Use AddFile + ConvertFiles for batch processing
'
' Output filename = DWGName_ViewName.pdf
'
' Prerequisites:
'   regsvr32 "path\to\dwgtopdfx64.dll"  (run as Administrator)
' ============================================================

Option Explicit

Dim obj, fso, folder, files, objFile
Set obj = CreateObject("DWGTOPDFX.ConvertPDF")
Set fso = CreateObject("Scripting.FileSystemObject")

' --- Configuration ---
Set folder = fso.GetFolder("D:\input\dwg")
Set files = folder.Files

WScript.Echo "Begin converting folder: " & folder.Name

' ============================================================
' Approach A: Convert each DWG file individually (per-view)
' ============================================================
Dim I, strViewName, strOutput

On Error Resume Next
For Each objFile In files
    If LCase(fso.GetExtensionName(objFile.Path)) <> "dwg" Then
        ' Skip non-DWG files
    Else
        obj.Addexcudelayer "Defpoints"
        obj.Width = 1024
        obj.Height = 768
        obj.ColorMode = 1
        obj.InputFile = objFile.Path
        obj.PenWidth(7) = 0.5
        obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"
        
        ' Convert each view in the DWG file
        For I = 1 To obj.ViewCount
            strViewName = obj.ViewName(I)
            strOutput = objFile.Name & "_" & strViewName & ".pdf"
            obj.ConvertView strOutput, strViewName
        Next
        
        If Err.Number <> 0 Then
            WScript.Echo "  Error converting " & objFile.Name & " (Error: 0x" & Hex(Err.Number) & ")"
            Err.Clear
        Else
            WScript.Echo "  Converted: " & objFile.Name & " (" & obj.ViewCount & " views)"
        End If
    End If
Next
On Error GoTo 0

' ============================================================
' Approach B: Use AddFile + ConvertFiles (batch mode)
' Uncomment to use:
' ============================================================
' WScript.Echo "Using AddFile + ConvertFiles approach..."
' Dim dwgFile
' For Each objFile In files
'     If LCase(fso.GetExtensionName(objFile.Path)) = "dwg" Then
'         obj.AddFile objFile.Path
'     End If
' Next
' obj.Width = 1024
' obj.Height = 768
' obj.ColorMode = 1
' obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"
' obj.ConvertFiles "D:\output\pdf\batch_output.pdf"

WScript.Echo "Batch conversion complete."

Set files = Nothing
Set folder = Nothing
Set fso = Nothing
Set obj = Nothing