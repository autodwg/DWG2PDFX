' Create conversion object for DWG to PDF conversion
Dim obj
Set obj = CreateObject("DWGTOPDFX.ConvertPDF")

' Configure non-printable layer conversion behavior
' 0 = Do not convert non-printable layers (default)
' 1 = Convert non-printable layers
obj.ConvertNotPlotLayer = 0

' Set output dimensions in millimeters
obj.Width  = 914    ' Page width (914mm ≈ 36 inches)
obj.Height = 610   ' Page height (610mm ≈ 24 inches)

' Configure color output mode
obj.ColorMode = 1  ' 1 = Black & White (0 = 256-color mode)

' Specify input DWG file
obj.InputFile = "Office.dwg"

' Set output resolution
OBJ.DPI = 100      ' Dots Per Inch setting

' Font configuration
obj.OutWithTTF = 1   ' Output TrueType fonts as text (not curves)
obj.EmbeddedSearchText = 0  ' Disable embedded searchable text

' Zoom configuration
obj.ZoomType = 0    ' 0 = Zoom to paper size (fit to page)

' Color mapping configuration
obj.SetCTBFile "acad.ctb"   ' Use "acad.ctb" color book for mapping

' Set output scale ratio
obj.SetOutputScale "1"     ' 1:1 scale (original size)

' Configure line thickness for color index 1 (red)
obj.PenWidth(1) = 0.02     ' Line width in mm (0.02mm pen)

' Configure PostScript support paths
obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts" ' PS resource paths

' Add SHX font directory to search path
obj.AddFontPath("C:\windows\system32\shxfont\")  ' Add SHX font support path

' User notifications
MsgBox "Begin converting..."  ' Start conversion prompt
obj.ConvertPapers("conveted.pdf")  ' Execute conversion
MsgBox "Conversion complete"  ' Completion notification

' Error handling
If Err.Number < 0 Then
    MsgBox Err.Description  ' Show error details if any
End If


	