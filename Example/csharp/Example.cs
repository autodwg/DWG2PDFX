// ============================================================
// AutoDWG DWG2PDF SDK - C# Examples
//
// Prerequisites:
//   1. Register the COM DLL (run as Administrator):
//      64-bit: regsvr32 "path\to\dwgtopdfx64.dll"
//      32-bit: regsvr32 "path\to\dwgtopdfx.dll"
//   2. .NET Framework 4.0+ or .NET 5+
//
// Option A - Late binding (no interop DLL needed):
//   csc /out:Example.exe Example.cs
//
// Option B - Early binding (add COM reference in Visual Studio):
//   Project > Add Reference > COM > DWGTOPDFX
//   Then use: DWGTOPDFXLib.ConvertPDFClass
// ============================================================

using System;
using System.IO;
using System.Runtime.InteropServices;

namespace AutoDWG.Examples
{
    class DWG2PDFExamples
    {
        static void Main(string[] args)
        {
            Console.WriteLine("=== AutoDWG DWG2PDF SDK - C# Examples ===\n");

            // Uncomment the example you want to run:
            Example1_BasicConversion();
            // Example2_OnePdfPerView();
            // Example3_BatchConversion(@"D:\input\dwg", @"D:\output\pdf");
            // Example4_AdvancedFeatures();
            // Example5_DWFtoDWG();
        }

        /// <summary>
        /// Example 1: Basic single file conversion.
        /// Converts all layouts of a DWG file into a single PDF.
        /// Demonstrates all basic properties including background color
        /// and separate DPI width/height.
        /// </summary>
        static void Example1_BasicConversion()
        {
            string inputFile = @"D:\input\dwg\Office.dwg";
            string outputFile = @"D:\output\pdf\output.pdf";

            Directory.CreateDirectory(Path.GetDirectoryName(outputFile));

            Type comType = Type.GetTypeFromProgID("DWGTOPDFX.ConvertPDF");
            if (comType == null)
            {
                Console.WriteLine("ERROR: COM object not found. Run: regsvr32 dwgtopdfx64.dll");
                return;
            }
            dynamic obj = Activator.CreateInstance(comType);

            try
            {
                obj.ConvertNotPlotLayer = 0;      // 0=Skip non-printing layers, 1=Convert them
                obj.Width = 240;                  // Page width in mm
                obj.Height = 180;                 // Page height in mm
                obj.ColorMode = 0;                // 0=256 colors, 1=Black & White
                obj.DPI = 300;                    // Output resolution (uniform)
                // obj.DPIWidth = 300;            // (Optional) Separate horizontal DPI
                // obj.DPIHeight = 300;           // (Optional) Separate vertical DPI
                obj.OutWithTTF = 1;               // 1=Preserve TrueType fonts
                obj.EmbeddedSearchText = 0;       // 0=Keep text searchable
                obj.ZoomType = 0;                 // 0=Zoom extents, 1=Last view
                obj.OutputByLayoutSize = 0;       // 0=Use Width/Height, 1=Use layout dims
                obj.nBackgroundColor = 7;  // Background color (7=white, 0=black)

                obj.set_PenWidth(1, 0.02);        // Pen width for color index 1 (Red)

                obj.AddFontPath(@"C:\windows\system32\shxfont");
                obj.PSPath = @"c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts";

                obj.InputFile = inputFile;

                Console.WriteLine("Converting: " + inputFile);
                obj.Convert(outputFile);
                Console.WriteLine("Done: " + outputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Conversion failed. (Error: 0x" + (ex.HResult & 0xFFFFFFFF).ToString("X8") + ")");
            }
            finally
            {
                Marshal.ReleaseComObject(obj);
            }
        }

        /// <summary>
        /// Example 2: Convert each layout/view to a separate PDF.
        /// Demonstrates SetOutColor for color index mapping and
        /// Addexcudelayer for excluding layers.
        /// </summary>
        static void Example2_OnePdfPerView()
        {
            string inputFile = @"D:\input\dwg\Office.dwg";
            string outputDir = @"D:\output\pdf";
            Directory.CreateDirectory(outputDir);

            Type comType = Type.GetTypeFromProgID("DWGTOPDFX.ConvertPDF");
            dynamic obj = Activator.CreateInstance(comType);

            try
            {
                obj.Addexcudelayer("Defpoints");
                obj.ConvertNotPlotLayer = 0;
                obj.Width = 300;
                obj.Height = 200;
                obj.ColorMode = 1;
                obj.PSPath = @"c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts";
                obj.InputFile = inputFile;

                for (int i = 1; i <= 7; i++)
                    obj.set_PenWidth(i, 0.04);

                // Color mapping: map input color index to output color
                // SetOutColor(inputColorIndex, outputColorIndex)
                obj.SetOutColor(1, 5);            // Map red to blue
                // obj.SetOutColor(2, 3);         // Map yellow to green
                // obj.SetOutColor(7, 1);         // Map white/black to red

                int viewCount = obj.ViewCount;
                Console.WriteLine("Found " + viewCount + " view(s).");

                for (int i = 1; i <= viewCount; i++)
                {
                    string viewName = obj.ViewName(i);
                    string pageSetup = obj.PageSetupName(i);
                    int unit = obj.PageSizeUnit(i);
                    double w = obj.PageWidth(i);
                    double h = obj.PageHeight(i);

                    Console.WriteLine("  [" + i + "] View: " + viewName +
                        ", Setup: " + pageSetup + ", Unit: " + unit +
                        ", Size: " + w + "x" + h + "mm");

                    string outFile = Path.Combine(outputDir, viewName + ".pdf");
                    obj.ConvertView(outFile, viewName);
                }

                Console.WriteLine("All views converted.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Conversion failed. (Error: 0x" + (ex.HResult & 0xFFFFFFFF).ToString("X8") + ")");
            }
            finally
            {
                Marshal.ReleaseComObject(obj);
            }
        }

        /// <summary>
        /// Example 3: Batch conversion of all DWG files in a folder.
        /// Approach A: Iterate folder, convert each DWG with all views.
        /// Approach B: Use AddFile + ConvertFiles (commented out).
        /// </summary>
        static void Example3_BatchConversion(string inputDir, string outputDir)
        {
            Directory.CreateDirectory(outputDir);
            string[] dwgFiles = Directory.GetFiles(inputDir, "*.dwg");
            if (dwgFiles.Length == 0)
            {
                Console.WriteLine("No DWG files found in: " + inputDir);
                return;
            }

            Console.WriteLine("Found " + dwgFiles.Length + " DWG file(s).");

            Type comType = Type.GetTypeFromProgID("DWGTOPDFX.ConvertPDF");
            dynamic obj = Activator.CreateInstance(comType);

            try
            {
                // --- Approach A: Convert each file individually ---
                foreach (string dwgFile in dwgFiles)
                {
                    obj.Addexcudelayer("Defpoints");
                    obj.Width = 1024;
                    obj.Height = 768;
                    obj.ColorMode = 1;
                    obj.InputFile = dwgFile;
                    obj.set_PenWidth(7, 0.5);
                    obj.PSPath = @"c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts";

                    int viewCount = obj.ViewCount;
                    string baseName = Path.GetFileNameWithoutExtension(dwgFile);

                    for (int i = 1; i <= viewCount; i++)
                    {
                        string viewName = obj.ViewName(i);
                        string outFile = Path.Combine(outputDir, baseName + "_" + viewName + ".pdf");
                        obj.ConvertView(outFile, viewName);
                    }

                    Console.WriteLine("  Converted: " + Path.GetFileName(dwgFile) + " (" + viewCount + " views)");
                }

                // --- Approach B: Use AddFile + ConvertFiles ---
                // foreach (string dwgFile in dwgFiles)
                //     obj.AddFile(dwgFile);
                // obj.Width = 1024;
                // obj.Height = 768;
                // obj.ColorMode = 1;
                // obj.PSPath = @"c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts";
                // obj.ConvertFiles(Path.Combine(outputDir, "batch_output.pdf"));

                Console.WriteLine("Batch conversion complete.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Conversion failed. (Error: 0x" + (ex.HResult & 0xFFFFFFFF).ToString("X8") + ")");
            }
            finally
            {
                Marshal.ReleaseComObject(obj);
            }
        }

        /// <summary>
        /// Example 4: Advanced features - CTB, scale, ConvertPapers, background color.
        /// </summary>
        static void Example4_AdvancedFeatures()
        {
            string inputFile = @"D:\input\dwg\Office.dwg";
            string outputFile = @"D:\output\pdf\output_papers.pdf";
            Directory.CreateDirectory(Path.GetDirectoryName(outputFile));

            Type comType = Type.GetTypeFromProgID("DWGTOPDFX.ConvertPDF");
            dynamic obj = Activator.CreateInstance(comType);

            try
            {
                obj.ConvertNotPlotLayer = 0;
                obj.Width = 914;
                obj.Height = 610;
                obj.ColorMode = 1;
                obj.DPI = 100;
                obj.OutWithTTF = 1;
                obj.EmbeddedSearchText = 0;
                obj.ZoomType = 0;
                obj.nBackgroundColor = 7;  // White background

                obj.SetCTBFile("acad.ctb");       // CTB color book
                obj.SetOutputScale(1.0f);         // 1:1 scale (float parameter)

                obj.set_PenWidth(1, 0.02);        // Red
                obj.set_PenWidth(2, 0.03);        // Yellow
                obj.set_PenWidth(3, 0.05);        // Green
                obj.set_PenWidth(4, 0.08);        // Cyan
                obj.set_PenWidth(5, 0.10);        // Blue
                obj.set_PenWidth(6, 0.15);        // Magenta
                obj.set_PenWidth(7, 0.20);        // White/Black

                obj.PSPath = @"c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts";
                obj.AddFontPath(@"C:\windows\system32\shxfont");
                obj.InputFile = inputFile;

                Console.WriteLine("Converting (papers only, no Model Space)...");
                obj.ConvertPapers(outputFile);
                Console.WriteLine("Done: " + outputFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Conversion failed. (Error: 0x" + (ex.HResult & 0xFFFFFFFF).ToString("X8") + ")");
            }
            finally
            {
                Marshal.ReleaseComObject(obj);
            }
        }

        /// <summary>
        /// Example 5: DWF to DWG/DXF conversion.
        /// Uses ConvertDWFtoDWGDXF method.
        /// </summary>
        static void Example5_DWFtoDWG()
        {
            string inputFile = @"D:\input\dwf\drawing.dwf";
            string outputFile = @"D:\output\dwg\output.dwg";
            short outVersion = 24;                // Output DWG version code

            Directory.CreateDirectory(Path.GetDirectoryName(outputFile));

            Type comType = Type.GetTypeFromProgID("DWGTOPDFX.ConvertPDF");
            dynamic obj = Activator.CreateInstance(comType);

            try
            {
                Console.WriteLine("Converting DWF to DWG...");
                Console.WriteLine("  Input:  " + inputFile);
                Console.WriteLine("  Output: " + outputFile);

                obj.ConvertDWFtoDWGDXF(inputFile, outputFile, outVersion);
                Console.WriteLine("Done: " + outputFile);

                // To convert to DXF instead:
                // string dxfOutput = Path.ChangeExtension(outputFile, ".dxf");
                // obj.ConvertDWFtoDWGDXF(inputFile, dxfOutput, outVersion);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Conversion failed. (Error: 0x" + (ex.HResult & 0xFFFFFFFF).ToString("X8") + ")");
            }
            finally
            {
                Marshal.ReleaseComObject(obj);
            }
        }
    }
}