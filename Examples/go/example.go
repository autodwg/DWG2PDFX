// ============================================================
// AutoDWG DWG2PDF SDK - Go Examples (using go-ole)
//
// Prerequisites:
//   1. Register the COM DLL (run as Administrator):
//      64-bit: regsvr32 "path\to\dwgtopdfx64.dll"
//      32-bit: regsvr32 "path\to\dwgtopdfx.dll"
//   2. Install go-ole:
//      go get github.com/go-ole/go-ole
//      go get github.com/go-ole/go-ole/oleutil
//
// Run:
//   go run example.go
// ============================================================

package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

var (
	inputDir  = `D:\input\dwg`
	outputDir = `D:\output\pdf`
)

func createConverter() *ole.IDispatch {
	unknown, err := oleutil.CreateObject("DWGTOPDFX.ConvertPDF")
	if err != nil {
		fmt.Println("ERROR: Cannot create COM object.")
		fmt.Println("Register DLL first: regsvr32 dwgtopdfx64.dll")
		log.Fatal(err)
	}
	obj, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Fatal(err)
	}
	return obj
}

func setProp(obj *ole.IDispatch, name string, value interface{}) {
	oleutil.PutProperty(obj, name, value)
}

func callMethod(obj *ole.IDispatch, name string, args ...interface{}) {
	oleutil.CallMethod(obj, name, args...)
}

// ============================================================
// Example 1: Basic single file conversion
// ============================================================
func example1_BasicConversion() {
	inputFile := filepath.Join(inputDir, "Office.dwg")
	outputFile := filepath.Join(outputDir, "output.pdf")
	os.MkdirAll(outputDir, 0755)

	obj := createConverter()
	defer obj.Release()

	setProp(obj, "ConvertNotPlotLayer", 0)
	setProp(obj, "Width", 240)
	setProp(obj, "Height", 180)
	setProp(obj, "ColorMode", 0)
	setProp(obj, "DPI", 300)
	setProp(obj, "OutWithTTF", 1)
	setProp(obj, "EmbeddedSearchText", 0)
	setProp(obj, "zoomtype", 0)
	setProp(obj, "OutputByLayoutSize", 0)
	callMethod(obj, "AddFontPath", `C:\windows\system32\shxfont`)
	setProp(obj, "PSPath", `c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts`)
	setProp(obj, "InputFile", inputFile)

	fmt.Println("Converting:", inputFile)
	callMethod(obj, "Convert", outputFile)
	fmt.Println("Done:", outputFile)
}

// ============================================================
// Example 2: One PDF per layout/view
// ============================================================
func example2_OnePdfPerView() {
	inputFile := filepath.Join(inputDir, "Office.dwg")
	os.MkdirAll(outputDir, 0755)

	obj := createConverter()
	defer obj.Release()

	callMethod(obj, "Addexcudelayer", "Defpoints")
	setProp(obj, "ConvertNotPlotLayer", 0)
	setProp(obj, "Width", 300)
	setProp(obj, "Height", 200)
	setProp(obj, "ColorMode", 1)
	setProp(obj, "PSPath", `c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts`)
	setProp(obj, "InputFile", inputFile)

	viewCountVar, _ := oleutil.GetProperty(obj, "ViewCount")
	viewCount := int(viewCountVar.Val)
	fmt.Printf("Found %d view(s).\n", viewCount)

	for i := 1; i <= viewCount; i++ {
		viewNameVar, _ := oleutil.CallMethod(obj, "ViewName", i)
		viewName := viewNameVar.ToString()

		pageSetupVar, _ := oleutil.CallMethod(obj, "PageSetupName", i)
		pageSetup := pageSetupVar.ToString()

		unitVar, _ := oleutil.CallMethod(obj, "PageSizeUnit", i)
		unit := int(unitVar.Val)

		wVar, _ := oleutil.CallMethod(obj, "PageWidth", i)
		w := wVar.Val

		hVar, _ := oleutil.CallMethod(obj, "PageHeight", i)
		h := hVar.Val

		fmt.Printf("  [%d] View: %s, Setup: %s, Unit: %d, Size: %vx%vmm\n",
			i, viewName, pageSetup, unit, w, h)

		outFile := filepath.Join(outputDir, viewName+".pdf")
		callMethod(obj, "ConvertView", outFile, viewName)
	}
	fmt.Println("All views converted.")
}

// ============================================================
// Example 3: Batch conversion
// ============================================================
func example3_BatchConversion() {
	os.MkdirAll(outputDir, 0755)

	entries, _ := os.ReadDir(inputDir)
	var dwgFiles []string
	for _, e := range entries {
		if strings.HasSuffix(strings.ToLower(e.Name()), ".dwg") {
			dwgFiles = append(dwgFiles, e.Name())
		}
	}
	if len(dwgFiles) == 0 {
		fmt.Println("No DWG files found.")
		return
	}
	fmt.Printf("Found %d DWG file(s).\n", len(dwgFiles))

	obj := createConverter()
	defer obj.Release()

	for _, dwgFile := range dwgFiles {
		fullPath := filepath.Join(inputDir, dwgFile)
		callMethod(obj, "Addexcudelayer", "Defpoints")
		setProp(obj, "Width", 1024)
		setProp(obj, "Height", 768)
		setProp(obj, "ColorMode", 1)
		setProp(obj, "InputFile", fullPath)
		setProp(obj, "PSPath", `c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts`)

		viewCountVar, _ := oleutil.GetProperty(obj, "ViewCount")
		viewCount := int(viewCountVar.Val)
		baseName := strings.TrimSuffix(dwgFile, filepath.Ext(dwgFile))

		for i := 1; i <= viewCount; i++ {
			viewNameVar, _ := oleutil.CallMethod(obj, "ViewName", i)
			viewName := viewNameVar.ToString()
			outFile := filepath.Join(outputDir, baseName+"_"+viewName+".pdf")
			callMethod(obj, "ConvertView", outFile, viewName)
		}
		fmt.Printf("  Converted: %s (%d views)\n", dwgFile, viewCount)
	}
	fmt.Println("Batch conversion complete.")
}

// ============================================================
// Example 4: Advanced features (CTB, scale, ConvertPapers)
// ============================================================
func example4_Advanced() {
	inputFile := filepath.Join(inputDir, "Office.dwg")
	outputFile := filepath.Join(outputDir, "output_papers.pdf")
	os.MkdirAll(outputDir, 0755)

	obj := createConverter()
	defer obj.Release()

	setProp(obj, "ConvertNotPlotLayer", 0)
	setProp(obj, "Width", 914)
	setProp(obj, "Height", 610)
	setProp(obj, "ColorMode", 1)
	setProp(obj, "DPI", 100)
	setProp(obj, "OutWithTTF", 1)
	setProp(obj, "EmbeddedSearchText", 0)
	setProp(obj, "ZoomType", 0)

	callMethod(obj, "SetCTBFile", "acad.ctb")
	callMethod(obj, "SetOutputScale", "1")

	setProp(obj, "PSPath", `c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts`)
	callMethod(obj, "AddFontPath", `C:\windows\system32\shxfont`)
	setProp(obj, "InputFile", inputFile)

	fmt.Println("Converting (papers only, no Model Space)...")
	callMethod(obj, "ConvertPapers", outputFile)
	fmt.Println("Done:", outputFile)
}

func main() {
	// Initialize COM
	if err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED); err != nil {
		ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED)
	}
	defer ole.CoUninitialize()

	fmt.Println("=== AutoDWG DWG2PDF SDK - Go Examples ===")
	fmt.Println()

	example1_BasicConversion()
	// example2_OnePdfPerView()
	// example3_BatchConversion()
	// example4_Advanced()

	fmt.Println("\nDone.")
}