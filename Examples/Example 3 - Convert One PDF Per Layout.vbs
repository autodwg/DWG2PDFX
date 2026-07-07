
' Create PDF conversion object using DWGTOPDFX COM component
Dim obj
Set obj = CreateObject("DWGTOPDFX.ConvertPDF")

' Configure layer settings
obj.Addexcudelayer "Defpoints"    ' Exclude non-printable layer (Defpoints)
obj.ConvertNotPlotLayer = 0       ' 0 = Do not convert non-printable layers (default), 1 = Convert them

' Set output dimensions in millimeters
obj.Width = 300                   ' PDF width (300mm)
obj.Height = 200                  ' PDF height (200mm)
obj.InputFile = "office.dwg"      ' Input DWG file path

' Configure pen widths for different line types (all set to 0.04mm)
obj.PenWidth(1) = 0.04            ' Line width for type 1
obj.PenWidth(2) = 0.04            ' Line width for type 2
obj.PenWidth(3) = 0.04            ' Line width for type 3
obj.PenWidth(4) = 0.04            ' Line width for type 4
obj.PenWidth(5) = 0.04            ' Line width for type 5
obj.PenWidth(6) = 0.04            ' Line width for type 6
obj.PenWidth(7) = 0.04            ' Line width for type 7

' Set color mode and background
obj.ColorMode = 1                 ' 1 = Black & White mode, 0 = 256-color mode
' obj.nBackgroundColor = 0        ' Default white background (0=black, 1=red...)
' Note: In monochrome mode (ColorMode=1), background is always white

' Configure PostScript paths for font/rendering support
obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts" 
' Additional font paths can be added with:
' obj.AddFontPath("C:\\Xref\\")
' obj.AddFontPath "E:\Program Files\AutoCAD 2007\Fonts"

' Display total views count message
MsgBox "Total " & obj.ViewCount & " Views"

' Process each view in the DWG file
For I = 1 to obj.ViewCount
    ' Get current view name
    strViewName = obj.ViewName(I)
    MsgBox "Converting view: " & strViewName & "..."
    
    ' Set output PDF path (same directory as script)
    strOutput = strViewName & ".pdf"  ' Output file name with path
    
    ' Display page setup details for debugging
    MsgBox "Page Name:" + obj.PageSetupName(I) + _
           " Unit:" + CStr(obj.PageSizeUnit(I)) + _
           " Width=" + CStr(obj.PageWidth(I)) + _
           " Height=" + CStr(obj.PageHeight(I))
    
    ' Perform conversion for current view
    obj.ConvertView strOutput, strViewName
Next

' Error handling section
If Err.Number < 0 Then
    MsgBox Err.Description  ' Show error details if any
End If


'Unit Value List
'UnitsUndefined    = 0,
'UnitsInches       = 1,
'UnitsFeet         = 2,
'UnitsMiles        = 3,
'UnitsMillimeters  = 4,
'UnitsCentimeters  = 5,
'UnitsMeters       = 6,
'UnitsKilometers   = 7,
'UnitsMicroinches  = 8,
'UnitsMils         = 9,
'UnitsYards        = 10,
'UnitsAngstroms    = 11,
'UnitsNanometers   = 12,
'UnitsMicrons      = 13,
'UnitsDecimeters   = 14,
'UnitsDekameters   = 15,
'UnitsHectometers  = 16,
'UnitsGigameters   = 17,
'UnitsAstronomical = 18,
'UnitsLightYears   = 19,
'UnitsParsecs      = 20,
'UnitsMax          = kUnitsParsecs	

