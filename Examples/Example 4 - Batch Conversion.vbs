' Create PDF conversion object using DWGTOPDFX COM component
Dim obj
Set obj = CreateObject("DWGTOPDFX.ConvertPDF")

' Initialize filesystem object for folder operations
Set fsoFolder = CreateObject("Scripting.FileSystemObject")
' Get reference to source folder (modify path as needed)
Set folder = fsoFolder.GetFolder("D:\AutoDWG\")
Set files = folder.Files  ' Get collection of files in folder

' Display start message with folder name
MsgBox "Begin to Convert folder: " & folder.Name

' Process each file in the folder
For Each objfile In files
    ' Configure layer exclusion settings
    obj.Addexcudelayer "Defpoints"  ' Exclude non-printable layer (Defpoints)
    
    ' Set output dimensions in pixels (1024x768 resolution)
    obj.Width = 1024
    obj.Height = 768
    
    ' Set color mode to black & white (7 corresponds to black/white in CAD index)
    obj.ColorMode = 1  ' 0-255: Color index, 7=black/white mode
    
    ' Specify input DWG file (current file in loop)
    obj.InputFile = objFile
    
    ' Configure pen width for black/white lines (AutoCAD color index 7)
    obj.PenWidth(7) = 0.5  ' Line width in pixels
    
    ' Set PostScript support paths for font/rendering
    obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"
    
    ' Additional font paths can be added with (commented examples):
    ' obj.AddFontPath("C:\\Xref\\")
    ' obj.AddFontPath "E:\Program Files\AutoCAD 2007\Fonts"
    
    ' Process each view in the current DWG file
    For I = 1 to obj.ViewCount
        strViewName = obj.ViewName(I)  ' Get view name
        ' Construct output filename (original filename + viewname + .pdf)
        strOutput = objFile.Name & strViewName & ".pdf"
        
        ' Perform view conversion
        obj.ConvertView strOutput, strViewName
    Next
    
    ' Error handling for current file conversion
    If Err.Number < 0 Then
        MsgBox Err.Description  ' Show error details
    End If
Next

' Display completion message
MsgBox "Conversion Complete"

' Cleanup resources
Set files = Nothing
Set folder = Nothing
Set fsoFolder = Nothing
