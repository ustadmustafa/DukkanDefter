using System.Globalization;
using DukkanDefterOCR.Models;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Extensions.Options;

namespace DukkanDefterOCR.Services
{
    public class GoogleSheetsService
    {
        private const string DevredenLabel = "Devreden akbil";
        private const string SatilanLabel = "Satılan Akbil";

        private readonly GoogleSheetsOptions _options;

        public GoogleSheetsService(IOptions<GoogleSheetsOptions> options)
        {
            _options = options.Value;
        }

        public async Task SaveToSheetAsync(string sheetName, IList<OCRItem> items, int devredenAkbil, int akbil, CancellationToken ct = default)
        {
            if (string.IsNullOrWhiteSpace(_options.CredentialsPath))
                throw new InvalidOperationException("GoogleSheets:CredentialsPath boş.");
            if (string.IsNullOrWhiteSpace(_options.SpreadsheetId))
                throw new InvalidOperationException("GoogleSheets:SpreadsheetId boş.");
            if (!File.Exists(_options.CredentialsPath))
                throw new FileNotFoundException("Credentials dosyası bulunamadı.", _options.CredentialsPath);

            await using var stream = new FileStream(_options.CredentialsPath, FileMode.Open, FileAccess.Read, FileShare.Read);
#pragma warning disable CS0618 // GoogleCredential.FromStream — servis hesabı JSON; CredentialFactory geçişi sonraki adım
            var credential = GoogleCredential
                .FromStream(stream)
                .CreateScoped(SheetsService.Scope.Spreadsheets);
#pragma warning restore CS0618

            using var service = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "DukkanDefterOCR"
            });

            var safeTitle = SanitizeSheetTitle(sheetName);

            var spreadsheet = await service.Spreadsheets.Get(_options.SpreadsheetId).ExecuteAsync(ct);
            var exists = spreadsheet.Sheets?.Any(s => s.Properties.Title == safeTitle) ?? false;

            if (!exists)
            {
                var batch = new BatchUpdateSpreadsheetRequest
                {
                    Requests = new List<Request>
                    {
                        new Request
                        {
                            AddSheet = new AddSheetRequest
                            {
                                Properties = new SheetProperties { Title = safeTitle }
                            }
                        }
                    }
                };
                await service.Spreadsheets.BatchUpdate(batch, _options.SpreadsheetId).ExecuteAsync(ct);
            }

            var mainValues = new List<IList<object>>
            {
                new List<object> { "Ürün", "Tutar" }
            };

            foreach (var item in items)
                mainValues.Add(new List<object> { item.Kalem, item.Toplam });

            var mainRowCount = mainValues.Count;
            var footerStartRow = mainRowCount + 3;

            var previousDevreden = await TryReadDevredenAkbilFromPreviousDayAsync(service, sheetName, ct);
            var satilanAkbil = akbil + (previousDevreden - devredenAkbil);

            var footerValues = new List<IList<object>>
            {
                new List<object> { DevredenLabel, devredenAkbil },
                new List<object> { SatilanLabel, satilanAkbil }
            };

            var escaped = safeTitle.Replace("'", "''", StringComparison.Ordinal);
            var vo = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            var mainBody = new ValueRange { Values = mainValues };
            var mainReq = service.Spreadsheets.Values.Update(mainBody, _options.SpreadsheetId, $"'{escaped}'!A1:B{mainRowCount}");
            mainReq.ValueInputOption = vo;
            await mainReq.ExecuteAsync(ct);

            var footerBody = new ValueRange { Values = footerValues };
            var footerReq = service.Spreadsheets.Values.Update(footerBody, _options.SpreadsheetId, $"'{escaped}'!C{footerStartRow}:D{footerStartRow + 1}");
            footerReq.ValueInputOption = vo;
            await footerReq.ExecuteAsync(ct);
        }

        private async Task<int> TryReadDevredenAkbilFromPreviousDayAsync(SheetsService service, string sheetName, CancellationToken ct)
        {
            if (!TryParseLedgerSheetDate(sheetName.Trim(), out var currentDate))
                return 0;

            var prevTitle = SanitizeSheetTitle(currentDate.AddDays(-1).ToString("dd.MM.yy", CultureInfo.InvariantCulture));
            var escapedPrev = prevTitle.Replace("'", "''", StringComparison.Ordinal);

            try
            {
                var range = $"'{escapedPrev}'!C1:D500";
                var req = service.Spreadsheets.Values.Get(_options.SpreadsheetId, range);
                var res = await req.ExecuteAsync(ct);
                if (res.Values == null)
                    return 0;

                foreach (var row in res.Values)
                {
                    if (row.Count < 1)
                        continue;
                    var label = row[0]?.ToString()?.Trim();
                    if (string.Equals(label, DevredenLabel, StringComparison.OrdinalIgnoreCase))
                    {
                        if (row.Count >= 2 && TryParseIntCell(row[1], out var v))
                            return v;
                        return 0;
                    }
                }

                return 0;
            }
            catch
            {
                return 0;
            }
        }

        private static bool TryParseLedgerSheetDate(string name, out DateTime date)
        {
            var formats = new[] { "dd.MM.yy", "dd.MM.yyyy" };
            foreach (var f in formats)
            {
                if (DateTime.TryParseExact(name, f, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                    return true;
            }

            date = default;
            return false;
        }

        private static bool TryParseIntCell(object? cell, out int value)
        {
            value = 0;
            if (cell == null)
                return false;

            switch (cell)
            {
                case int i:
                    value = i;
                    return true;
                case long l:
                    value = (int)l;
                    return true;
                case double d:
                    value = (int)d;
                    return true;
                case decimal m:
                    value = (int)m;
                    return true;
            }

            var s = cell.ToString()?.Trim();
            if (string.IsNullOrEmpty(s))
                return false;

            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var iv))
            {
                value = iv;
                return true;
            }

            if (decimal.TryParse(s, NumberStyles.Number, CultureInfo.GetCultureInfo("tr-TR"), out var tr))
            {
                value = (int)tr;
                return true;
            }

            return false;
        }

        private static string SanitizeSheetTitle(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Tarihsiz";

            foreach (var c in new[] { '\\', '/', '?', '*', '[', ']' })
                name = name.Replace(c, '_');

            name = name.Trim();
            if (name.Length > 99)
                name = name[..99];

            return name;
        }
    }
}
