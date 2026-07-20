#nullable enable
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Http;
using AutoDWG.DWG2PDF.Web.Services;

namespace AutoDWG.DWG2PDF.Web.Pages
{
    public class IndexModel : PageModel
    {
        private readonly DWGConverterService _converter;

        public IndexModel(DWGConverterService converter)
        {
            _converter = converter;
        }

        [BindProperty]
        public IFormFile? DwgFile { get; set; }

        [BindProperty]
        public int ColorMode { get; set; } = 0;

        [BindProperty]
        public int DPI { get; set; } = 300;

        [BindProperty]
        public int Width { get; set; } = 297;

        [BindProperty]
        public int Height { get; set; } = 210;

        [BindProperty]
        public bool OutWithTTF { get; set; } = true;

        [BindProperty]
        public bool EmbeddedSearchText { get; set; }

        [BindProperty]
        public int ZoomType { get; set; } = 0;

        [BindProperty]
        public bool OutputByLayoutSize { get; set; }

        [BindProperty]
        public int BackgroundColor { get; set; } = 7;

        [BindProperty]
        public string FontPath { get; set; } = @"C:\windows\system32\shxfont";

        [BindProperty]
        public string PSPath { get; set; } = @"c:\windows\system32\ps\lib;c:\windows\system32\ps\fonts";

        public ConversionResult? Result { get; set; }

        public void OnGet() { }

        public IActionResult OnPost()
        {
            if (DwgFile == null || DwgFile.Length == 0)
            {
                Result = new ConversionResult { Success = false, Message = "Please select a DWG file." };
                return Page();
            }

            // Save uploaded file
            var uploadPath = Path.Combine(_converter.UploadDir, DwgFile.FileName);
            using (var stream = new FileStream(uploadPath, FileMode.Create))
            {
                DwgFile.CopyTo(stream);
            }

            // Build conversion request
            var request = new ConversionRequest
            {
                ColorMode = ColorMode,
                DPI = DPI,
                Width = Width,
                Height = Height,
                OutWithTTF = OutWithTTF,
                EmbeddedSearchText = EmbeddedSearchText,
                ZoomType = ZoomType,
                OutputByLayoutSize = OutputByLayoutSize,
                BackgroundColor = BackgroundColor,
                FontPath = FontPath,
                PSPath = PSPath
            };

            // Convert
            Result = _converter.Convert(uploadPath, request);
            return Page();
        }

        // Download converted PDF
        public IActionResult OnGetDownload(string file)
        {
            var filePath = Path.Combine(_converter.OutputDir, file);
            if (!System.IO.File.Exists(filePath))
                return NotFound();

            var memoryStream = new MemoryStream(System.IO.File.ReadAllBytes(filePath));
            return File(memoryStream, "application/pdf", file);
        }
    }
}