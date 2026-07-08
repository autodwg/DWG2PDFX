# DWG2PDFX
**AutoDWG DWG to PDF Component**

## AutoDWG DWG to PDF Control Component

DWG2PDF-X, a control component allows you to convert dwg to pdf, dxf and dwf to pdf directly, no AutoCAD required.

### Key features:
- Support DWG, DXF and DWF in versions from R2.5 to 2027.
- Create PDF file with or without Model space.
- Create an individual PDF file per layout.
- Support all standard and customize output paper size.
- Batch mode supported.

### Free Trial Download Link
https://github.com/autodwg/DWG2PDFX/releases/download/v1.9.8/DWG2PDFX-v1.9.8.zip

## Getting Started
### Quick setup
 (The steps below are for 64-bit installation.)

#### Step 1: Register the DLL Component

Double-click 'reg.bat' to automatically register 'DWGTOPDFX64.dll' on your system.

If registration fails: 
Open Command Prompt as Administrator via:
Start Menu → Windows System → Right-click "Command Prompt" → Run as Administrator
Manually register the DLL using: 
```cmd
regsvr32 DWGTOPDFX64.dll
```
#### Step 2: Configure PostScript (PS) Support

Copy PS Files: Move all PS-related files/folders to the default directory: C:\Windows\System32\PS\

Update Registry: Run PSpath64.reg to automatically add the PS path to the system registry.
Manual Registry Edit (Alternative): Navigate to: HKEY_LOCAL_MACHINE\SOFTWARE\AFPL Ghostscript\8.53

Verify/create the GS_LIB key with the following value:
C:\Windows\System32\ps\lib;C:\Windows\System32\ps\fonts

#### Step 3: Test with Example VBScript
Use the provided sample script (example.vbs) in the Examples folder to verify functionality.
Ensure the script executes without errors and generates the expected PDF output.

### Sample code
#### Convert all layouts to one pdf
AutoDWG DWG2PDFX allows you convert all layouts to one pdf.

VBS Code
```
Dim obj
Set obj= CreateObject("DWGTOPDFX.ConvertPDF")
obj.AddExcudeLayer "Defpoints"							
obj.ConvertNotPlotLayer = 0	
obj.Width  = 914							
obj.Height = 610								
obj.ColorMode = 0								
obj.InputFile = "sample.dwg"
obj.DPI = 100
obj.zoomtype=0
obj.PenWidth(1) = 0.02	
obj.PSPath =  "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"				
obj.AddFontPath("C:\\Fonts\\")			
MsgBox "Begin converting..."
obj.Convert "sample.pdf"
MsgBox "Conversion complete"
If Err.Number < 0 Then
MsgBox Err.Description
End If
```
## License Notice
1. Free trial / non-commercial use: GNU LGPLv2.1
2. Commercial production use, closed-source integration requires purchasing our commercial license.

Contact info@autodwg.com for commercial authorization.
