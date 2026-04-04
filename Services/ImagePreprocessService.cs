using Microsoft.AspNetCore.Hosting;
using OpenCvSharp;

namespace DukkanDefterOCR.Services
{
    /// <summary>
    /// Defter fotoğrafları için OCR öncesi kontrast / ikili görüntü iyileştirmesi.
    /// </summary>
    public class ImagePreprocessService
    {
        public string CreateProcessedImagePath(string sourcePath, IWebHostEnvironment env)
        {
            var uploads = Path.Combine(env.WebRootPath, "uploads");
            Directory.CreateDirectory(uploads);
            var outName = "proc_" + Guid.NewGuid().ToString("N") + ".png";
            var outPath = Path.Combine(uploads, outName);

            using var src = Cv2.ImRead(sourcePath, ImreadModes.Color);
            if (src.Empty())
                throw new InvalidOperationException("Görüntü okunamadı: " + sourcePath);

            using var work = ResizeIfHuge(src);
            using var gray = new Mat();
            Cv2.CvtColor(work, gray, ColorConversionCodes.BGR2GRAY);

            using var blur = new Mat();
            Cv2.GaussianBlur(gray, blur, new Size(5, 5), 0);

            using var thresh = new Mat();
            Cv2.AdaptiveThreshold(
                blur,
                thresh,
                255,
                AdaptiveThresholdTypes.GaussianC,
                ThresholdTypes.Binary,
                15,
                10);

            Cv2.ImWrite(outPath, thresh);
            return outPath;
        }

        private static Mat ResizeIfHuge(Mat src)
        {
            const int maxDim = 2200;
            var w = src.Width;
            var h = src.Height;
            if (w <= maxDim && h <= maxDim)
                return src.Clone();

            var scale = maxDim / (double)Math.Max(w, h);
            var nw = (int)(w * scale);
            var nh = (int)(h * scale);
            var dst = new Mat();
            Cv2.Resize(src, dst, new Size(nw, nh), 0, 0, InterpolationFlags.Area);
            return dst;
        }
    }
}
