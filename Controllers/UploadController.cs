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

        public async Task<IActionResult> Index(CancellationToken ct)
        {
            var id = _sheetsOptions.SpreadsheetId?.Trim();
            if (!string.IsNullOrEmpty(id))
                ViewData["GoogleSpreadsheetOpenUrl"] = $"https://docs.google.com/spreadsheets/d/{id}/edit";

            var vm = new ManualEntryViewModel { SheetDate = DateTime.Today.ToString("dd.MM.yy") };
            for (var i = 0; i < 5; i++)
                vm.Rows.Add(new ManualSheetRow());

            try
            {
                var (devir, dunKasa) = await _googleSheetsService.TryReadPreviousDayClosingAsync(vm.SheetDate, ct);
                vm.DevirAlinanAkbilPreview = devir;
                vm.DunKasaFromPreviousDay = dunKasa;
            }
            catch
            {
                // Kimlik dosyası yoksa veya ağ hatası: form yine açılır.
            }

            return View(vm);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Save(ManualEntryViewModel model, CancellationToken ct)
        {
            if (!model.KasayaGirilenPara.HasValue)
            {
                TempData["Error"] = "Kasaya girilen para girin.";
                return RedirectToAction(nameof(Index));
            }

            if (!model.Kasa.HasValue)
            {
                TempData["Error"] = "Kasa tutarını girin.";
                return RedirectToAction(nameof(Index));
            }

            if (!model.YuklenenAkbil.HasValue)
            {
                TempData["Error"] = "Yüklenen akbil girin.";
                return RedirectToAction(nameof(Index));
            }

            if (!model.DevredecekAkbil.HasValue)
            {
                TempData["Error"] = "Devredecek akbil girin.";
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

            try
            {
                var sheetDate = string.IsNullOrWhiteSpace(model.SheetDate) ? "Tarihsiz" : model.SheetDate.Trim();
                var (devirAlinan, _) = await _googleSheetsService.TryReadPreviousDayClosingAsync(sheetDate, ct);

                await _googleSheetsService.SaveToSheetAsync(
                    sheetDate,
                    items,
                    model.KasayaGirilenPara.Value,
                    model.Kasa.Value,
                    model.YuklenenAkbil.Value,
                    model.DevredecekAkbil.Value,
                    devirAlinan,
                    ct);

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
