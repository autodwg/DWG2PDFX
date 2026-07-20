// ============================================================
// AutoDWG DWG2PDF SDK - ASP.NET Core Web Application
//
// Prerequisites:
//   1. Register the COM DLL (run as Administrator):
//      64-bit: regsvr32 "path\to\dwgtopdfx64.dll"
//   2. .NET 8.0 SDK
//
// Run:
//   dotnet run
//   Opens automatically at http://localhost:9876
// ============================================================

using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using AutoDWG.DWG2PDF.Web.Services;

var builder = WebApplication.CreateBuilder(args);

// Configure port 9876
builder.WebHost.UseUrls("http://localhost:9876");

// Add services
builder.Services.AddRazorPages();
builder.Services.AddSingleton<DWGConverterService>();

var app = builder.Build();

// Configure middleware
app.UseStaticFiles();
app.UseRouting();
app.MapRazorPages();

// Auto-open browser after startup
var lifetime = app.Services.GetRequiredService<IHostApplicationLifetime>();
lifetime.ApplicationStarted.Register(() =>
{
    Task.Delay(500).ContinueWith(_ =>
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "http://localhost:9876",
                UseShellExecute = true
            });
            Console.WriteLine("Browser opened at http://localhost:9876");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Could not open browser: {ex.Message}");
        }
    });
});

Console.WriteLine("==============================================");
Console.WriteLine("  AutoDWG DWG to PDF - Web Converter");
Console.WriteLine("  Running at: http://localhost:9876");
Console.WriteLine("==============================================");

app.Run();

// ============================================================
// Service classes (must be in a namespace)
// ============================================================
namespace AutoDWG.DWG2PDF.Web.Services
{
    public class ConversionRequest
    {
        public int ColorMode { get; set; }
        public int DPI { get; set; } = 300;
        public int Width { get; set; } = 297;
        public int Height { get; set; } = 210;
        public bool OutWithTTF { get; set; } = true;
        public bool EmbeddedSearchText { get; set; }
        public int ZoomType { get; set; }
        public bool OutputByLayoutSize { get; set; }
        public int BackgroundColor { get; set; } = 7; // White
        public string FontPath { get; set; } = "";
        public string PSPath { get; set; } = "";
    }

    public class ConversionResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = "";
        public string OutputPath { get; set; } = "";
        public string OutputFileName { get; set; } = "";
    }

    public class DWGConverterService
    {
        private readonly string _uploadDir;
        private readonly string _outputDir;

        public DWGConverterService()
        {
            _uploadDir = Path.Combine(AppContext.BaseDirectory, "uploads");
            _outputDir = Path.Combine(AppContext.BaseDirectory, "output");
            Directory.CreateDirectory(_uploadDir);
            Directory.CreateDirectory(_outputDir);
        }

        public string UploadDir => _uploadDir;
        public string OutputDir => _outputDir;

        public ConversionResult Convert(string inputFilePath, ConversionRequest request)
        {
            var result = new ConversionResult();
            var outputFileName = Path.GetFileNameWithoutExtension(inputFilePath) + ".pdf";
            var outputPath = Path.Combine(_outputDir, outputFileName);

            try
            {
                Type? comType = Type.GetTypeFromProgID("DWGTOPDFX.ConvertPDF");
                if (comType == null)
                    throw new Exception("COM object not found. Run: regsvr32 dwgtopdfx64.dll");

                dynamic obj = Activator.CreateInstance(comType!);
                try
                {
                    obj.ColorMode = request.ColorMode;
                    obj.DPI = request.DPI;
                    obj.OutWithTTF = request.OutWithTTF ? 1 : 0;
                    obj.EmbeddedSearchText = request.EmbeddedSearchText ? 1 : 0;
                    obj.ZoomType = request.ZoomType;
                    obj.OutputByLayoutSize = request.OutputByLayoutSize ? 1 : 0;
                    obj.Width = request.Width;
                    obj.Height = request.Height;
                    obj.nBackgroundColor = request.BackgroundColor;

                    if (!string.IsNullOrWhiteSpace(request.FontPath))
                        obj.AddFontPath(request.FontPath);
                    if (!string.IsNullOrWhiteSpace(request.PSPath))
                        obj.PSPath = request.PSPath;

                    obj.InputFile = inputFilePath;
                    obj.Convert(outputPath);

                    result.Success = true;
                    result.Message = "Conversion completed successfully.";
                    result.OutputPath = outputPath;
                    result.OutputFileName = outputFileName;
                }
                finally
                {
                    if (obj != null)
                        Marshal.ReleaseComObject(obj);
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = "Conversion failed. Please check the input file and settings, then try again. (Error: 0x" + (ex.HResult & 0xFFFFFFFF).ToString("X8") + ")";
            }

            return result;
        }
    }
}