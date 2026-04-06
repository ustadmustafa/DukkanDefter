using DukkanDefterOCR.Models;
using DukkanDefterOCR.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace DukkanDefterOCR.Controllers
{
    public class UploadController : Controller
    {
        private readonly GoogleSheetsService _googleSheetsService;
        private readonly GoogleSheetsOptions _sheetsOptions;

        public UploadController(GoogleSheetsService googleSheetsService, IOptions<GoogleSheetsOptions> sheetsOptions)
        {
            _googleSheetsService = googleSheetsService;
            _sheetsOptions = sheetsOptions.Value;
        }

        public IActionResult Index()
        {
            var id = _sheetsOptions.SpreadsheetId?.Trim();
            if (!string.IsNullOrEmpty(id))
                ViewData["GoogleSpreadsheetOpenUrl"] = $"https://docs.google.com/spreadsheets/d/{id}/edit";

            var vm = new ManualEntryViewModel { SheetDate = DateTime.Today.ToString("dd.MM.yy") };
            for (var i = 0; i < 5; i++)
                vm.Rows.Add(new ManualSheetRow());
            return View(vm);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Save(ManualEntryViewModel model, CancellationToken ct)
        {
            if (!model.DevredenAkbil.HasValue)
            {
                TempData["Error"] = "Devreden Akbil girin.";
                return RedirectToAction(nameof(Index));
            }

            if (model.Rows == null || model.Rows.Count == 0)
            {
                TempData["Error"] = "En az bir satır gerekli.";
                return RedirectToAction(nameof(Index));
            }

            var filled = model.Rows
                .Where(r => !string.IsNullOrWhiteSpace(r.Urun))
                .ToList();

            if (filled.Count == 0)
            {
                TempData["Error"] = "En az bir satırda ürün adı yazın.";
                return RedirectToAction(nameof(Index));
            }

            var items = filled.Select(r => new OCRItem
            {
                Kalem = r.Urun.Trim(),
                Toplam = r.Tutar,
                HamTutar = "",
                Not = "",
                IsGrandTotalRow = false
            }).ToList();

            var sum = items.Sum(i => i.Toplam);
            items.Add(new OCRItem
            {
                Kalem = "Toplam",
                HamTutar = "",
                Toplam = sum,
                Not = "",
                IsGrandTotalRow = true
            });

            try
            {
                var sheetDate = string.IsNullOrWhiteSpace(model.SheetDate) ? "Tarihsiz" : model.SheetDate.Trim();
                await _googleSheetsService.SaveToSheetAsync(sheetDate, items, model.DevredenAkbil.Value, ct);
                TempData["Success"] = $"'{sheetDate}' sekmesine kaydedildi.";
            }
            catch (Exception ex)
            {
                TempData["Error"] = ex.Message;
            }

            return RedirectToAction(nameof(Index));
        }
    }
}
