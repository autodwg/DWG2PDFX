' Create an instance of the DWG-to-PDF converter object
Dim obj
Set obj = CreateObject("DWGTOPDFX.ConvertPDF")

' Configure non-printing layer conversion: 0=Do not convert non-printing layers (default), 1=Convert non-printing layers
obj.ConvertNotPlotLayer = 0

' Set output PDF dimensions (Units: mm)
obj.Width  = 240   ' Page width
obj.Height = 180   ' Page height

' Color mode configuration: 0=256-color mode, 1=Black & White mode
obj.ColorMode = 0

' Specify input DWG file path
obj.InputFile = "office.dwg"

' Set output resolution (DPI)
obj.DPI = 300

' Font output option: 1=Preserve TrueType fonts (prevents converting to graphics)
obj.OutWithTTF = 1

' Embed searchable text: 0=Keep text searchable, 1=Embed as image (prevents text selection)
obj.EmbeddedSearchText = 0

' Set default pen width (Units: mm, here 0.02mm pen width)
obj.PenWidth(1) = 0.02

' View scaling mode: 0=Zoom extents (show full drawing), 1=Use last view state
obj.zoomtype = 0

' Layout size output option: 0=Use specified dimensions, 1=Use layout's original dimensions
obj.OutputByLayoutSize = 0

' Configure PostScript support files path (Requires pre-copied PS files to these locations)
obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"

' Add SHX font support path (AutoCAD vector fonts)
obj.AddFontPath("C:\windows\system32\shxfont")

' Pre-conversion notification
MsgBox "Begin converting..."

' Execute conversion and specify output path
obj.Convert "output.pdf"

' Post-conversion completion notification
MsgBox "Conversion complete"

' Error handling (Capture conversion process exceptions)
If Err.Number < 0 Then
    MsgBox "Error: " & Err.Description
End If