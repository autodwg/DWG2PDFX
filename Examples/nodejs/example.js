// ============================================================
// AutoDWG DWG2PDF SDK - Node.js Examples (using win32-ole)
//
// Prerequisites:
//   1. Register the COM DLL (run as Administrator):
//      64-bit: regsvr32 "path\to\dwgtopdfx64.dll"
//      32-bit: regsvr32 "path\to\dwgtopdfx.dll"
//   2. Install dependencies:
//      npm install
//
// Run:
//   node example.js
// ============================================================

const path = require("path");
const fs = require("fs");

const inputDir = "D:\\input\\dwg";
const outputDir = "D:\\output\\pdf";
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

// --- Load COM library ---
let ole;
try {
    ole = require("win32-ole");
} catch (e) {
    try {
        ole = require("node-win32com");
    } catch (e2) {
        console.error("ERROR: Install a COM library:");
        console.error("  npm install win32-ole");
        console.error("  npm install node-win32com");
        process.exit(1);
    }
}

function createConverter() {
    try {
        return new ole.Client("DWGTOPDFX.ConvertPDF");
    } catch (e) {
        console.error("ERROR: Cannot create COM object.");
        console.error("Register DLL first: regsvr32 dwgtopdfx64.dll");
        process.exit(1);
    }
}

// ============================================================
// Example 1: Basic single file conversion
// ============================================================
function example1_BasicConversion() {
    const inputFile = path.join(inputDir, "Office.dwg");
    const outputFile = path.join(outputDir, "output.pdf");

    const obj = createConverter();
    try {
        obj.Set("ConvertNotPlotLayer", 0);
        obj.Set("Width", 240);  obj.Set("Height", 180);
        obj.Set("ColorMode", 0);  obj.Set("DPI", 300);
        obj.Set("OutWithTTF", 1);  obj.Set("EmbeddedSearchText", 0);
        obj.Set("zoomtype", 0);  obj.Set("OutputByLayoutSize", 0);
        obj.Call("AddFontPath", "C:\\windows\\system32\\shxfont");
        obj.Set("PSPath", "c:\\windows\\system32\\ps\\lib;c:\\windows\\system32\\ps\\fonts");
        obj.Set("InputFile", inputFile);

        console.log("Converting:", inputFile);
        obj.Call("Convert", outputFile);
        console.log("Done:", outputFile);
    } catch (e) {
        console.error("Failed:", e.message);
    } finally {
        obj.Close();
    }
}

// ============================================================
// Example 2: One PDF per layout/view
// ============================================================
function example2_OnePdfPerView() {
    const inputFile = path.join(inputDir, "Office.dwg");

    const obj = createConverter();
    try {
        obj.Call("Addexcudelayer", "Defpoints");
        obj.Set("ConvertNotPlotLayer", 0);
        obj.Set("Width", 300);  obj.Set("Height", 200);
        obj.Set("ColorMode", 1);
        obj.Set("PSPath", "c:\\windows\\system32\\ps\\lib;c:\\windows\\system32\\ps\\fonts");
        obj.Set("InputFile", inputFile);

        const viewCount = obj.Get("ViewCount");
        console.log(`Found ${viewCount} view(s).`);

        for (let i = 1; i <= viewCount; i++) {
            const viewName = obj.Call("ViewName", i);
            const pageSetup = obj.Call("PageSetupName", i);
            const unit = obj.Call("PageSizeUnit", i);
            const w = obj.Call("PageWidth", i);
            const h = obj.Call("PageHeight", i);

            console.log(`  [${i}] View: ${viewName}, Setup: ${pageSetup}, Unit: ${unit}, Size: ${w}x${h}mm`);

            const outFile = path.join(outputDir, `${viewName}.pdf`);
            obj.Call("ConvertView", outFile, viewName);
        }
        console.log("All views converted.");
    } catch (e) {
        console.error("Failed:", e.message);
    } finally {
        obj.Close();
    }
}

// ============================================================
// Example 3: Batch conversion
// ============================================================
function example3_BatchConversion() {
    const dwgFiles = fs.readdirSync(inputDir).filter(f => f.toLowerCase().endsWith(".dwg"));
    if (!dwgFiles.length) {
        console.log("No DWG files found.");
        return;
    }
    console.log(`Found ${dwgFiles.length} DWG file(s).`);

    const obj = createConverter();
    try {
        for (const dwgFile of dwgFiles) {
            const fullPath = path.join(inputDir, dwgFile);
            obj.Call("Addexcudelayer", "Defpoints");
            obj.Set("Width", 1024);  obj.Set("Height", 768);
            obj.Set("ColorMode", 1);
            obj.Set("InputFile", fullPath);
            obj.Set("PSPath", "c:\\windows\\system32\\ps\\lib;c:\\windows\\system32\\ps\\fonts");

            const viewCount = obj.Get("ViewCount");
            const baseName = path.basename(dwgFile, ".dwg");

            for (let i = 1; i <= viewCount; i++) {
                const viewName = obj.Call("ViewName", i);
                const outFile = path.join(outputDir, `${baseName}_${viewName}.pdf`);
                obj.Call("ConvertView", outFile, viewName);
            }
            console.log(`  Converted: ${dwgFile} (${viewCount} views)`);
        }
        console.log("Batch conversion complete.");
    } catch (e) {
        console.error("Failed:", e.message);
    } finally {
        obj.Close();
    }
}

// ============================================================
// Example 4: Advanced features
// ============================================================
function example4_Advanced() {
    const inputFile = path.join(inputDir, "Office.dwg");
    const outputFile = path.join(outputDir, "output_papers.pdf");

    const obj = createConverter();
    try {
        obj.Set("ConvertNotPlotLayer", 0);
        obj.Set("Width", 914);  obj.Set("Height", 610);
        obj.Set("ColorMode", 1);  obj.Set("DPI", 100);
        obj.Set("OutWithTTF", 1);  obj.Set("EmbeddedSearchText", 0);  obj.Set("ZoomType", 0);

        obj.Call("SetCTBFile", "acad.ctb");
        obj.Call("SetOutputScale", "1");

        obj.Set("PSPath", "c:\\windows\\system32\\ps\\lib;c:\\windows\\system32\\ps\\fonts");
        obj.Call("AddFontPath", "C:\\windows\\system32\\shxfont");
        obj.Set("InputFile", inputFile);

        console.log("Converting (papers only, no Model Space)...");
        obj.Call("ConvertPapers", outputFile);
        console.log("Done:", outputFile);
    } catch (e) {
        console.error("Failed:", e.message);
    } finally {
        obj.Close();
    }
}

// --- Run Example 1 ---
console.log("=== AutoDWG DWG2PDF SDK - Node.js Examples ===\n");
example1_BasicConversion();
// example2_OnePdfPerView();
// example3_BatchConversion();
// example4_Advanced();
console.log("\nDone.");