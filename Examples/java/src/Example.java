// ============================================================
// AutoDWG DWG2PDF SDK - Java Examples (using Jacob)
//
// Prerequisites:
//   1. Register the COM DLL (run as Administrator):
//      64-bit: regsvr32 "path\to\dwgtopdfx64.dll"
//      32-bit: regsvr32 "path\to\dwgtopdfx.dll"
//   2. Download Jacob: https://github.com/freemansoft/jacob-project
//      - Place jacob.jar in lib/
//      - Place jacob-x64.dll in lib/
//
// Build & Run:
//   javac -cp "lib\jacob.jar" -d out src\Example.java
//   java -cp "out;lib\jacob.jar" -Djava.library.path=lib Example
// ============================================================

package com.autodwg.example;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;

public class Example {

    private static ActiveXComponent createConverter() {
        try {
            return new ActiveXComponent("DWGTOPDFX.ConvertPDF");
        } catch (Exception e) {
            System.err.println("ERROR: Cannot create COM object.");
            System.err.println("Register DLL first: regsvr32 dwgtopdfx64.dll");
            System.exit(1);
            return null;
        }
    }

    // ============================================================
    // Example 1: Basic single file conversion
    // ============================================================
    public static void example1_BasicConversion() {
        String inputFile = "D:\\input\\dwg\\Office.dwg";
        String outputFile = "D:\\output\\pdf\\output.pdf";
        new File(outputFile).getParentFile().mkdirs();

        ComThread.InitMTA();
        ActiveXComponent obj = createConverter();
        try {
            Dispatch converter = obj.getObject();

            Dispatch.put(converter, "ConvertNotPlotLayer", new Variant(0));
            Dispatch.put(converter, "Width", new Variant(240));
            Dispatch.put(converter, "Height", new Variant(180));
            Dispatch.put(converter, "ColorMode", new Variant(0));     // 256 colors
            Dispatch.put(converter, "DPI", new Variant(300));
            Dispatch.put(converter, "OutWithTTF", new Variant(1));
            Dispatch.put(converter, "EmbeddedSearchText", new Variant(0));
            Dispatch.put(converter, "zoomtype", new Variant(0));       // Zoom extents
            Dispatch.put(converter, "OutputByLayoutSize", new Variant(0));

            Dispatch.call(converter, "AddFontPath",
                new Variant("C:\\windows\\system32\\shxfont"));
            Dispatch.put(converter, "PSPath",
                new Variant("c:\\windows\\system32\\ps\\lib;c:\\windows\\system32\\ps\\fonts"));
            Dispatch.put(converter, "InputFile", new Variant(inputFile));

            System.out.println("Converting: " + inputFile);
            Dispatch.call(converter, "Convert", new Variant(outputFile));
            System.out.println("Done: " + outputFile);
        } catch (Exception e) {
            System.err.println("Failed: " + e.getMessage());
        } finally {
            obj.safeRelease();
            ComThread.Release();
        }
    }

    // ============================================================
    // Example 2: One PDF per layout/view
    // ============================================================
    public static void example2_OnePdfPerView() {
        String inputFile = "D:\\input\\dwg\\Office.dwg";
        String outputDir = "D:\\output\\pdf";
        new File(outputDir).mkdirs();

        ComThread.InitMTA();
        ActiveXComponent obj = createConverter();
        try {
            Dispatch converter = obj.getObject();

            Dispatch.call(converter, "Addexcudelayer", new Variant("Defpoints"));
            Dispatch.put(converter, "ConvertNotPlotLayer", new Variant(0));
            Dispatch.put(converter, "Width", new Variant(300));
            Dispatch.put(converter, "Height", new Variant(200));
            Dispatch.put(converter, "ColorMode", new Variant(1));      // B&W
            Dispatch.put(converter, "PSPath",
                new Variant("c:\\windows\\system32\\ps\\lib;c:\\windows\\system32\\ps\\fonts"));
            Dispatch.put(converter, "InputFile", new Variant(inputFile));

            int viewCount = Dispatch.get(converter, "ViewCount").toInt();
            System.out.println("Found " + viewCount + " view(s).");

            for (int i = 1; i <= viewCount; i++) {
                String viewName = Dispatch.call(converter, "ViewName", new Variant(i)).toString();
                String pageSetup = Dispatch.call(converter, "PageSetupName", new Variant(i)).toString();
                int unit = Dispatch.call(converter, "PageSizeUnit", new Variant(i)).toInt();
                double w = Dispatch.call(converter, "PageWidth", new Variant(i)).toDouble();
                double h = Dispatch.call(converter, "PageHeight", new Variant(i)).toDouble();

                System.out.printf("  [%d] View: %s, Setup: %s, Unit: %d, Size: %.0fx%.0fmm%n",
                    i, viewName, pageSetup, unit, w, h);

                String outFile = outputDir + File.separator + viewName + ".pdf";
                Dispatch.call(converter, "ConvertView",
                    new Variant(outFile), new Variant(viewName));
            }
            System.out.println("All views converted.");
        } catch (Exception e) {
            System.err.println("Failed: " + e.getMessage());
        } finally {
            obj.safeRelease();
            ComThread.Release();
        }
    }

    // ============================================================
    // Example 3: Batch conversion
    // ============================================================
    public static void example3_BatchConversion(String inputDir, String outputDir) {
        new File(outputDir).mkdirs();
        File dir = new File(inputDir);
        File[] dwgFiles = dir.listFiles((d, name) -> name.toLowerCase().endsWith(".dwg"));

        if (dwgFiles == null || dwgFiles.length == 0) {
            System.out.println("No DWG files found in: " + inputDir);
            return;
        }

        System.out.println("Found " + dwgFiles.length + " DWG file(s).");

        ComThread.InitMTA();
        ActiveXComponent obj = createConverter();
        try {
            Dispatch converter = obj.getObject();

            for (File dwgFile : dwgFiles) {
                Dispatch.call(converter, "Addexcudelayer", new Variant("Defpoints"));
                Dispatch.put(converter, "Width", new Variant(1024));
                Dispatch.put(converter, "Height", new Variant(768));
                Dispatch.put(converter, "ColorMode", new Variant(1));
                Dispatch.put(converter, "InputFile", new Variant(dwgFile.getAbsolutePath()));
                Dispatch.put(converter, "PSPath",
                    new Variant("c:\\windows\\system32\\ps\\lib;c:\\windows\\system32\\ps\\fonts"));

                int viewCount = Dispatch.get(converter, "ViewCount").toInt();
                String baseName = dwgFile.getName().replaceFirst("[.][^.]+$", "");

                for (int i = 1; i <= viewCount; i++) {
                    String viewName = Dispatch.call(converter, "ViewName", new Variant(i)).toString();
                    String outFile = outputDir + File.separator + baseName + "_" + viewName + ".pdf";
                    Dispatch.call(converter, "ConvertView",
                        new Variant(outFile), new Variant(viewName));
                }
                System.out.println("  Converted: " + dwgFile.getName() + " (" + viewCount + " views)");
            }
            System.out.println("Batch conversion complete.");
        } catch (Exception e) {
            System.err.println("Failed: " + e.getMessage());
        } finally {
            obj.safeRelease();
            ComThread.Release();
        }
    }

    // ============================================================
    // Example 4: Advanced features (CTB, scale, ConvertPapers)
    // ============================================================
    public static void example4_Advanced() {
        String inputFile = "D:\\input\\dwg\\Office.dwg";
        String outputFile = "D:\\output\\pdf\\output_papers.pdf";
        new File(outputFile).getParentFile().mkdirs();

        ComThread.InitMTA();
        ActiveXComponent obj = createConverter();
        try {
            Dispatch converter = obj.getObject();

            Dispatch.put(converter, "ConvertNotPlotLayer", new Variant(0));
            Dispatch.put(converter, "Width", new Variant(914));
            Dispatch.put(converter, "Height", new Variant(610));
            Dispatch.put(converter, "ColorMode", new Variant(1));
            Dispatch.put(converter, "DPI", new Variant(100));
            Dispatch.put(converter, "OutWithTTF", new Variant(1));
            Dispatch.put(converter, "EmbeddedSearchText", new Variant(0));
            Dispatch.put(converter, "ZoomType", new Variant(0));

            Dispatch.call(converter, "SetCTBFile", new Variant("acad.ctb"));
            Dispatch.call(converter, "SetOutputScale", new Variant("1"));

            Dispatch.put(converter, "PSPath",
                new Variant("c:\\windows\\system32\\ps\\lib;c:\\windows\\system32\\ps\\fonts"));
            Dispatch.call(converter, "AddFontPath",
                new Variant("C:\\windows\\system32\\shxfont"));
            Dispatch.put(converter, "InputFile", new Variant(inputFile));

            System.out.println("Converting (papers only, no Model Space)...");
            Dispatch.call(converter, "ConvertPapers", new Variant(outputFile));
            System.out.println("Done: " + outputFile);
        } catch (Exception e) {
            System.err.println("Failed: " + e.getMessage());
        } finally {
            obj.safeRelease();
            ComThread.Release();
        }
    }

    public static void main(String[] args) {
        System.out.println("=== AutoDWG DWG2PDF SDK - Java Examples ===\n");

        example1_BasicConversion();
        // example2_OnePdfPerView();
        // example3_BatchConversion("D:\\input\\dwg", "D:\\output\\pdf");
        // example4_Advanced();

        System.out.println("\nDone.");
    }
}