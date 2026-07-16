' ============================================================
' AutoDWG DWG2PDF SDK - Example 3: Batch Folder Conversion
'
' Converts all DWG files in a folder, each view to a
' separate PDF. Output filename = DWGName + ViewName.pdf
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

MsgBox "Begin converting folder: " & folder.Name

For Each objFile In files
    ' Only process .dwg files
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
        Dim I, strViewName, strOutput
        For I = 1 To obj.ViewCount
            strViewName = obj.ViewName(I)
            strOutput = objFile.Name & "_" & strViewName & ".pdf"
            obj.ConvertView strOutput, strViewName
        Next
        
        If Err.Number < 0 Then
            MsgBox "Error converting " & objFile.Name & ": " & Err.Description
            Err.Clear
        End If
    End If
Next

MsgBox "Batch conversion complete."

Set files = Nothing
Set folder = Nothing
Set fso = Nothing
Set obj = Nothing