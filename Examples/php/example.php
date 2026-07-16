<?php
// ============================================================
// AutoDWG DWG2PDF SDK - PHP Examples
//
// Prerequisites:
//   1. Register the COM DLL (run as Administrator):
//      64-bit: regsvr32 "path\to\dwgtopdfx64.dll"
//      32-bit: regsvr32 "path\to\dwgtopdfx.dll"
//   2. PHP on Windows with COM extension: extension=com_dotnet
//
// Run:
//   php example.php
// ============================================================

$inputDir  = "D:\\input\\dwg";
$outputDir = "D:\\output\\pdf";
if (!is_dir($outputDir)) mkdir($outputDir, 0755, true);

function createConverter() {
    try {
        return new COM("DWGTOPDFX.ConvertPDF");
    } catch (Exception $e) {
        echo "ERROR: Cannot create COM object.\n";
        echo "Register DLL first: regsvr32 dwgtopdfx64.dll\n";
        exit(1);
    }
}

// ============================================================
// Example 1: Basic single file conversion
// ============================================================
function example1_basic() {
    global $outputDir;
    $inputFile = "D:\\input\\dwg\\Office.dwg";
    $outputFile = $outputDir . "\\output.pdf";

    $obj = createConverter();
    try {
        $obj->ConvertNotPlotLayer = 0;
        $obj->Width = 240;  $obj->Height = 180;
        $obj->ColorMode = 0;  $obj->DPI = 300;
        $obj->OutWithTTF = 1;  $obj->EmbeddedSearchText = 0;
        $obj->zoomtype = 0;  $obj->OutputByLayoutSize = 0;
        $obj->AddFontPath("C:\\windows\\system32\\shxfont");
        $obj->PSPath = "c:\\windows\\system32\\ps\\lib;c:\\windows\\system32\\ps\\fonts";
        $obj->InputFile = $inputFile;

        echo "Converting: $inputFile\n";
        $obj->Convert($outputFile);
        echo "Done: $outputFile\n";
    } catch (Exception $e) {
        echo "Failed: " . $e->getMessage() . "\n";
    } finally {
        $obj->Release();
    }
}

// ============================================================
// Example 2: One PDF per layout/view
// ============================================================
function example2_onePdfPerView() {
    global $outputDir;
    $inputFile = "D:\\input\\dwg\\Office.dwg";

    $obj = createConverter();
    try {
        $obj->Addexcudelayer("Defpoints");
        $obj->ConvertNotPlotLayer = 0;
        $obj->Width = 300;  $obj->Height = 200;
        $obj->ColorMode = 1;
        $obj->PSPath = "c:\\windows\\system32\\ps\\lib;c:\\windows\\system32\\ps\\fonts";
        $obj->InputFile = $inputFile;

        $viewCount = $obj->ViewCount;
        echo "Found $viewCount view(s).\n";

        for ($i = 1; $i <= $viewCount; $i++) {
            $viewName = $obj->ViewName($i);
            $pageSetup = $obj->PageSetupName($i);
            $unit = $obj->PageSizeUnit($i);
            $w = $obj->PageWidth($i);
            $h = $obj->PageHeight($i);

            echo "  [$i] View: $viewName, Setup: $pageSetup, Unit: $unit, Size: {$w}x{$h}mm\n";

            $outFile = $outputDir . "\\" . $viewName . ".pdf";
            $obj->ConvertView($outFile, $viewName);
        }
        echo "All views converted.\n";
    } catch (Exception $e) {
        echo "Failed: " . $e->getMessage() . "\n";
    } finally {
        $obj->Release();
    }
}

// ============================================================
// Example 3: Batch conversion
// ============================================================
function example3_batchConversion() {
    global $inputDir, $outputDir;
    $dwgFiles = glob($inputDir . "\\*.dwg");

    if (!$dwgFiles) {
        echo "No DWG files found.\n";
        return;
    }
    echo "Found " . count($dwgFiles) . " DWG file(s).\n";

    $obj = createConverter();
    try {
        foreach ($dwgFiles as $dwgFile) {
            $obj->Addexcudelayer("Defpoints");
            $obj->Width = 1024;  $obj->Height = 768;
            $obj->ColorMode = 1;
            $obj->InputFile = $dwgFile;
            $obj->PSPath = "c:\\windows\\system32\\ps\\lib;c:\\windows\\system32\\ps\\fonts";

            $viewCount = $obj->ViewCount;
            $baseName = pathinfo($dwgFile, PATHINFO_FILENAME);

            for ($i = 1; $i <= $viewCount; $i++) {
                $viewName = $obj->ViewName($i);
                $outFile = $outputDir . "\\" . $baseName . "_" . $viewName . ".pdf";
                $obj->ConvertView($outFile, $viewName);
            }
            echo "  Converted: " . basename($dwgFile) . " ($viewCount views)\n";
        }
        echo "Batch conversion complete.\n";
    } catch (Exception $e) {
        echo "Failed: " . $e->getMessage() . "\n";
    } finally {
        $obj->Release();
    }
}

// ============================================================
// Example 4: Advanced features (CTB, scale, ConvertPapers)
// ============================================================
function example4_advanced() {
    global $outputDir;
    $inputFile = "D:\\input\\dwg\\Office.dwg";
    $outputFile = $outputDir . "\\output_papers.pdf";

    $obj = createConverter();
    try {
        $obj->ConvertNotPlotLayer = 0;
        $obj->Width = 914;  $obj->Height = 610;
        $obj->ColorMode = 1;  $obj->DPI = 100;
        $obj->OutWithTTF = 1;  $obj->EmbeddedSearchText = 0;  $obj->ZoomType = 0;

        $obj->SetCTBFile("acad.ctb");
        $obj->SetOutputScale("1");

        $obj->PSPath = "c:\\windows\\system32\\ps\\lib;c:\\windows\\system32\\ps\\fonts";
        $obj->AddFontPath("C:\\windows\\system32\\shxfont");
        $obj->InputFile = $inputFile;

        echo "Converting (papers only, no Model Space)...\n";
        $obj->ConvertPapers($outputFile);
        echo "Done: $outputFile\n";
    } catch (Exception $e) {
        echo "Failed: " . $e->getMessage() . "\n";
    } finally {
        $obj->Release();
    }
}

// --- Run Example 1 ---
echo "=== AutoDWG DWG2PDF SDK - PHP Examples ===\n\n";
example1_basic();
// example2_onePdfPerView();
// example3_batchConversion();
// example4_advanced();
echo "\nDone.\n";