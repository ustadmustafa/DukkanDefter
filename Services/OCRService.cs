using System.Runtime.InteropServices;
using Microsoft.AspNetCore.Hosting;
using Tesseract;

namespace DukkanDefterOCR.Services
{
    public class OCRService
    {
        private readonly IConfiguration _configuration;
        private readonly ImagePreprocessService _preprocess;
        private readonly IWebHostEnvironment _env;

        public OCRService(
            IConfiguration configuration,
            ImagePreprocessService preprocess,
            IWebHostEnvironment env)
        {
            _configuration = configuration;
            _preprocess = preprocess;
            _env = env;
        }

        public string ExtractText(string imagePath)
        {
            var tessdataPath = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? _configuration["Tesseract:WindowsTessdataPath"]
                : _configuration["Tesseract:LinuxTessdataPath"];

            if (string.IsNullOrWhiteSpace(tessdataPath))
                throw new InvalidOperationException("Tesseract tessdata yolu yapılandırılmamış.");

            var usePreprocess = _configuration.GetValue("Ocr:UsePreprocessing", true);
            var processedPath = imagePath;

            if (usePreprocess)
            {
                try
                {
                    processedPath = _preprocess.CreateProcessedImagePath(imagePath, _env);
                }
                catch
                {
                    processedPath = imagePath;
                }
            }

            try
            {
                using var engine = new TesseractEngine(tessdataPath, "tur+eng", EngineMode.Default);
                using var img = Pix.LoadFromFile(processedPath);
                using var page = engine.Process(img, PageSegMode.SingleColumn);

                return page.GetText();
            }
            finally
            {
                if (usePreprocess && processedPath != imagePath && File.Exists(processedPath))
                {
                    try { File.Delete(processedPath); } catch { /* ignore */ }
                }
            }
        }
    }
}
