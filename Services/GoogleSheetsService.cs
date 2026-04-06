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
        private readonly GoogleSheetsOptions _options;

        public GoogleSheetsService(IOptions<GoogleSheetsOptions> options)
        {
            _options = options.Value;
        }

        public async Task SaveToSheetAsync(string sheetName, IList<OCRItem> items, int devredenAkbil, CancellationToken ct = default)
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

            var service = new SheetsService(new BaseClientService.Initializer
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

            var values = new List<IList<object>>
            {
                new List<object> { "Ürün", "Tutar" }
            };

            foreach (var item in items)
            {
                values.Add(new List<object> { item.Kalem, item.Toplam });
            }

            values.Add(new List<object> { "Devreden akbil", devredenAkbil });

            var body = new ValueRange { Values = values };
            var escaped = safeTitle.Replace("'", "''", StringComparison.Ordinal);
            var range = $"'{escaped}'!A1";
            var request = service.Spreadsheets.Values.Update(body, _options.SpreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            await request.ExecuteAsync(ct);
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
