# ============================================================
# AutoDWG DWG2PDF SDK - PowerShell Examples
#
# Prerequisites:
#   1. Register the COM DLL (run as Administrator):
#      64-bit: regsvr32 "path\to\dwgtopdfx64.dll"
#      32-bit: regsvr32 "path\to\dwgtopdfx.dll"
#
# Run:
#   powershell -ExecutionPolicy Bypass -File .\example.ps1
# ============================================================

Write-Host "=== AutoDWG DWG2PDF SDK - PowerShell Examples ===" -ForegroundColor Cyan
Write-Host ""

# --- Configuration ---
$inputDir  = "D:\input\dwg"
$outputDir = "D:\output\pdf"
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null

# --- Create COM Object ---
try {
    $obj = New-Object -ComObject "DWGTOPDFX.ConvertPDF"
} catch {
    Write-Host "ERROR: Cannot create COM object." -ForegroundColor Red
    Write-Host "Please register: regsvr32 dwgtopdfx64.dll"
    exit 1
}

# ============================================================
# Example 1: Basic single file conversion
# ============================================================
function Convert-Basic {
    param([string]$InputFile = "$inputDir\Office.dwg",
          [string]$OutputFile = "$outputDir\output.pdf")

    $obj.ConvertNotPlotLayer = 0
    $obj.Width = 240;  $obj.Height = 180
    $obj.ColorMode = 0;  $obj.DPI = 300
    $obj.OutWithTTF = 1;  $obj.EmbeddedSearchText = 0
    $obj.zoomtype = 0;  $obj.OutputByLayoutSize = 0
    $obj.AddFontPath("C:\windows\system32\shxfont")
    $obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"
    $obj.InputFile = $InputFile

    Write-Host "Converting: $InputFile"
    try {
        $obj.Convert($OutputFile)
        Write-Host "Done: $OutputFile" -ForegroundColor Green
    } catch {
        Write-Host "Failed: $_" -ForegroundColor Red
    }
}

# ============================================================
# Example 2: One PDF per layout/view
# ============================================================
function Convert-ByView {
    param([string]$InputFile = "$inputDir\Office.dwg",
          [string]$OutDir = $outputDir)

    $obj.Addexcudelayer("Defpoints")
    $obj.ConvertNotPlotLayer = 0
    $obj.Width = 300;  $obj.Height = 200
    $obj.ColorMode = 1
    $obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"
    $obj.InputFile = $InputFile

    $viewCount = $obj.ViewCount
    Write-Host "Found $viewCount view(s)."

    for ($i = 1; $i -le $viewCount; $i++) {
        $viewName = $obj.ViewName($i)
        $pageSetup = $obj.PageSetupName($i)
        $unit = $obj.PageSizeUnit($i)
        $w = $obj.PageWidth($i)
        $h = $obj.PageHeight($i)

        Write-Host "  [$i] View: $viewName, Setup: $pageSetup, Unit: $unit, Size: ${w}x${h}mm"

        $outFile = Join-Path $OutDir "$viewName.pdf"
        $obj.ConvertView($outFile, $viewName)
    }
    Write-Host "All views converted." -ForegroundColor Green
}

# ============================================================
# Example 3: Batch conversion
# ============================================================
function Convert-Batch {
    param([string]$SrcDir = $inputDir, [string]$DstDir = $outputDir)

    $dwgFiles = Get-ChildItem -Path $SrcDir -Filter "*.dwg"
    if (-not $dwgFiles) {
        Write-Host "No DWG files found in: $SrcDir" -ForegroundColor Yellow
        return
    }

    Write-Host "Found $($dwgFiles.Count) DWG file(s)."

    foreach ($file in $dwgFiles) {
        $obj.Addexcudelayer("Defpoints")
        $obj.Width = 1024;  $obj.Height = 768
        $obj.ColorMode = 1
        $obj.InputFile = $file.FullName
        $obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"

        $viewCount = $obj.ViewCount
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

        for ($i = 1; $i -le $viewCount; $i++) {
            $viewName = $obj.ViewName($i)
            $outFile = Join-Path $DstDir "${baseName}_${viewName}.pdf"
            $obj.ConvertView($outFile, $viewName)
        }

        Write-Host "  Converted: $($file.Name) ($viewCount views)"
    }
    Write-Host "Batch conversion complete." -ForegroundColor Green
}

# ============================================================
# Example 4: Advanced features
# ============================================================
function Convert-Advanced {
    param([string]$InputFile = "$inputDir\Office.dwg",
          [string]$OutputFile = "$outputDir\output_papers.pdf")

    $obj.ConvertNotPlotLayer = 0
    $obj.Width = 914;  $obj.Height = 610
    $obj.ColorMode = 1;  $obj.DPI = 100
    $obj.OutWithTTF = 1;  $obj.EmbeddedSearchText = 0;  $obj.ZoomType = 0

    $obj.SetCTBFile("acad.ctb")
    $obj.SetOutputScale("1")

    $obj.PSPath = "c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts"
    $obj.AddFontPath("C:\windows\system32\shxfont")
    $obj.InputFile = $InputFile

    Write-Host "Converting (papers only, no Model Space)..."
    try {
        $obj.ConvertPapers($OutputFile)
        Write-Host "Done: $OutputFile" -ForegroundColor Green
    } catch {
        Write-Host "Failed: $_" -ForegroundColor Red
    }
}

# --- Run Example 1 (uncomment others as needed) ---
Convert-Basic
# Convert-ByView
# Convert-Batch
# Convert-Advanced

# --- Cleanup ---
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($obj) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "`nDone."