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
        private const string KasayaGirilenLabel = "Kasaya Girilen Para";
        private const string KasaLabel = "Kasa";
        private const string ToplamHarcamalarLabel = "Toplam Harcamalar";
        private const string DevredenAkbilLabel = "Devreden Akbil";
        private const string SatilanAkbilLabel = "Satılan Akbil";
        private const string YuklenenAkbilLabel = "Yüklenen Akbil";
        private const string GunSonuLabel = "Gün Sonu";
        private const string DevredecekAkbilLabel = "Devredecek Akbil";
        private const string DevirAlinanAkbilLabel = "Devir Alınan Akbil";
        private const string CiroLabel = "Ciro";

        /// <summary>Önceki gün sekmesine formülle referans için sabit satırlar (E:F).</summary>
        private const int MirrorRowKasaya = 1;
        private const int MirrorRowKasa = 2;
        private const int MirrorRowDevredecek = 3;
        private const int MirrorRowYuklenen = 4;
        private const int MirrorRowDevirAlinan = 5;

        private readonly GoogleSheetsOptions _options;

        public GoogleSheetsService(IOptions<GoogleSheetsOptions> options)
        {
            _options = options.Value;
        }

        /// <summary>Bugünkü sekmeye göre bir önceki günün <c>Devredecek Akbil</c> (bugünün devir alınanı) ve <c>Kasa</c> değerlerini okur.</summary>
        public async Task<(int devirAlinanAkbil, int dunKasa)> TryReadPreviousDayClosingAsync(string sheetName, CancellationToken ct = default)
        {
            if (string.IsNullOrWhiteSpace(_options.CredentialsPath) || string.IsNullOrWhiteSpace(_options.SpreadsheetId))
                return (0, 0);
            if (!File.Exists(_options.CredentialsPath))
                return (0, 0);

            await using var stream = new FileStream(_options.CredentialsPath, FileMode.Open, FileAccess.Read, FileShare.Read);
#pragma warning disable CS0618
            var credential = GoogleCredential.FromStream(stream).CreateScoped(SheetsService.Scope.Spreadsheets);
#pragma warning restore CS0618

            using var service = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "DukkanDefterOCR"
            });

            if (!TryParseLedgerSheetDate(sheetName.Trim(), out var currentDate))
                return (0, 0);

            var prevTitle = SanitizeSheetTitle(currentDate.AddDays(-1).ToString("dd.MM.yy", CultureInfo.InvariantCulture));
            var escapedPrev = prevTitle.Replace("'", "''", StringComparison.Ordinal);

            try
            {
                var mirrorRange = $"'{escapedPrev}'!E{MirrorRowKasaya}:F{MirrorRowDevirAlinan}";
                var req = service.Spreadsheets.Values.Get(_options.SpreadsheetId, mirrorRange);
                var res = await req.ExecuteAsync(ct);
                var devir = 0;
                var dunKasa = 0;
                if (res.Values != null)
                {
                    foreach (var row in res.Values)
                    {
                        if (row.Count < 2)
                            continue;
                        var label = row[0]?.ToString()?.Trim();
                        if (string.Equals(label, DevredecekAkbilLabel, StringComparison.OrdinalIgnoreCase) && TryParseIntCell(row[1], out var dv))
                            devir = dv;
                        if (string.Equals(label, KasaLabel, StringComparison.OrdinalIgnoreCase) && TryParseIntCell(row[1], out var kv))
                            dunKasa = kv;
                    }
                }

                if (devir == 0 || dunKasa == 0)
                {
                    var (dFallback, kFallback) = await TryReadLegacyPrevDayFromAbAsync(service, escapedPrev, ct);
                    if (devir == 0)
                        devir = dFallback;
                    if (dunKasa == 0)
                        dunKasa = kFallback;
                }

                return (devir, dunKasa);
            }
            catch
            {
                return (0, 0);
            }
        }

        private async Task<(int devir, int dunKasa)> TryReadLegacyPrevDayFromAbAsync(SheetsService service, string escapedPrevTitle, CancellationToken ct)
        {
            var devir = 0;
            var dunKasa = 0;
            try
            {
                var range = $"'{escapedPrevTitle}'!A1:B500";
                var req = service.Spreadsheets.Values.Get(_options.SpreadsheetId, range);
                var res = await req.ExecuteAsync(ct);
                if (res.Values == null)
                    return (0, 0);

                foreach (var row in res.Values)
                {
                    if (row.Count < 2)
                        continue;
                    var label = row[0]?.ToString()?.Trim();
                    if (string.Equals(label, DevredecekAkbilLabel, StringComparison.OrdinalIgnoreCase) && TryParseIntCell(row[1], out var dv))
                        devir = dv;
                    if (string.Equals(label, KasaLabel, StringComparison.OrdinalIgnoreCase) && TryParseIntCell(row[1], out var kv))
                        dunKasa = kv;
                }

                if (devir == 0)
                {
                    foreach (var row in res.Values)
                    {
                        if (row.Count < 2)
                            continue;
                        var label = row[0]?.ToString()?.Trim();
                        if (string.Equals(label, "Devreden akbil", StringComparison.OrdinalIgnoreCase) && TryParseIntCell(row[1], out var d2))
                        {
                            devir = d2;
                            break;
                        }
                    }
                }
            }
            catch
            {
                // yoksay
            }

            return (devir, dunKasa);
        }

        public async Task SaveToSheetAsync(
            string sheetName,
            IList<OCRItem> expenseItems,
            int kasayaGirilenPara,
            int kasa,
            int yuklenenAkbil,
            int devredecekAkbil,
            int devirAlinanAkbil,
            CancellationToken ct = default)
        {
            if (string.IsNullOrWhiteSpace(_options.CredentialsPath))
                throw new InvalidOperationException("GoogleSheets:CredentialsPath boş.");
            if (string.IsNullOrWhiteSpace(_options.SpreadsheetId))
                throw new InvalidOperationException("GoogleSheets:SpreadsheetId boş.");
            if (!File.Exists(_options.CredentialsPath))
                throw new FileNotFoundException("Credentials dosyası bulunamadı.", _options.CredentialsPath);

            await using var stream = new FileStream(_options.CredentialsPath, FileMode.Open, FileAccess.Read, FileShare.Read);
#pragma warning disable CS0618
            var credential = GoogleCredential.FromStream(stream).CreateScoped(SheetsService.Scope.Spreadsheets);
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

            var e = expenseItems.Count;
            var rExpEnd = 2 + e;
            var rKasa = 3 + e;
            var rYuklenen = 5 + e;
            var rDevredecek = 6 + e;
            var rDevirAlinan = 7 + e;
            var rToplamHarc = 9 + e;
            var rDevreden = 11 + e;
            var rSatilan = 13 + e;
            var rGunSonu = 14 + e;
            var rCiro = 15 + e;
            var mainRowCount = 15 + e;

            var mainValues = new List<IList<object>>
            {
                new List<object> { KasayaGirilenLabel, kasayaGirilenPara },
                new List<object> { "Ürün", "Tutar" }
            };

            foreach (var item in expenseItems)
                mainValues.Add(new List<object> { item.Kalem, item.Toplam });

            mainValues.Add(new List<object> { KasaLabel, kasa });
            mainValues.Add(new List<object> { "", "" });
            mainValues.Add(new List<object> { YuklenenAkbilLabel, yuklenenAkbil });
            mainValues.Add(new List<object> { DevredecekAkbilLabel, devredecekAkbil });
            mainValues.Add(new List<object> { DevirAlinanAkbilLabel, devirAlinanAkbil });
            mainValues.Add(new List<object> { "", "" });
            mainValues.Add(new List<object> { ToplamHarcamalarLabel, $"=SUM(B3:B{rExpEnd})+B{rKasa}-B1" });
            mainValues.Add(new List<object> { "", "" });
            mainValues.Add(new List<object> { DevredenAkbilLabel, devirAlinanAkbil });
            mainValues.Add(new List<object> { "", "" });
            mainValues.Add(new List<object> { SatilanAkbilLabel, $"=B{rDevredecek}+B{rYuklenen}-B{rDevirAlinan}" });
            mainValues.Add(new List<object> { GunSonuLabel, $"=SUM(B3:B{rExpEnd})+B{rYuklenen}-B1+B{rKasa}" });

            var prevRef = TryGetPreviousSheetFormulaRef(sheetName);
            var ciroFormula = prevRef != null
                ? $"=B{rGunSonu}-'{prevRef}'!$F${MirrorRowKasa}-B{rSatilan}"
                : $"=B{rGunSonu}-0-B{rSatilan}";
            mainValues.Add(new List<object> { CiroLabel, ciroFormula });

            var mirrorValues = new List<IList<object>>
            {
                new List<object> { KasayaGirilenLabel, kasayaGirilenPara },
                new List<object> { KasaLabel, kasa },
                new List<object> { DevredecekAkbilLabel, devredecekAkbil },
                new List<object> { YuklenenAkbilLabel, yuklenenAkbil },
                new List<object> { DevirAlinanAkbilLabel, devirAlinanAkbil }
            };

            var escaped = safeTitle.Replace("'", "''", StringComparison.Ordinal);
            var vo = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

            var mainBody = new ValueRange { Values = mainValues };
            var mainReq = service.Spreadsheets.Values.Update(mainBody, _options.SpreadsheetId, $"'{escaped}'!A1:B{mainRowCount}");
            mainReq.ValueInputOption = vo;
            await mainReq.ExecuteAsync(ct);

            var mirrorBody = new ValueRange { Values = mirrorValues };
            var mirrorReq = service.Spreadsheets.Values.Update(mirrorBody, _options.SpreadsheetId, $"'{escaped}'!E{MirrorRowKasaya}:F{MirrorRowDevirAlinan}");
            mirrorReq.ValueInputOption = vo;
            await mirrorReq.ExecuteAsync(ct);

            var clearLegacy = new List<IList<object>>();
            for (var i = 0; i < 45; i++)
                clearLegacy.Add(new List<object> { "", "" });
            var clearReq = service.Spreadsheets.Values.Update(new ValueRange { Values = clearLegacy }, _options.SpreadsheetId, $"'{escaped}'!C1:D45");
            clearReq.ValueInputOption = vo;
            await clearReq.ExecuteAsync(ct);

            spreadsheet = await service.Spreadsheets.Get(_options.SpreadsheetId).ExecuteAsync(ct);
            var sheet = spreadsheet.Sheets?.FirstOrDefault(s => s.Properties.Title == safeTitle);
            var sheetId = sheet?.Properties.SheetId;
            if (sheetId == null)
                return;

            var formatRequests = BuildLedgerFormatRequests(sheetId.Value, e, mainRowCount);
            if (formatRequests.Count > 0)
            {
                await service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = formatRequests }, _options.SpreadsheetId).ExecuteAsync(ct);
            }
        }

        private static string? TryGetPreviousSheetFormulaRef(string sheetName)
        {
            if (!TryParseLedgerSheetDate(sheetName.Trim(), out var currentDate))
                return null;

            var prevTitle = SanitizeSheetTitle(currentDate.AddDays(-1).ToString("dd.MM.yy", CultureInfo.InvariantCulture));
            return prevTitle.Replace("'", "''", StringComparison.Ordinal);
        }

        private static List<Request> BuildLedgerFormatRequests(int sheetId, int expenseCount, int mainRowCount1Based)
        {
            var requests = new List<Request>();
            var e = expenseCount;
            var kasaRow0 = 2 + e;
            var yuklenenRow0 = 4 + e;
            var devredecekRow0 = 5 + e;
            var devirRow0 = 6 + e;
            var toplamRow0 = 8 + e;
            var devredenRow0 = 10 + e;
            var satilanRow0 = 12 + e;
            var gunsonuRow0 = 13 + e;
            var ciroRow0 = 14 + e;
            var mainEnd0 = mainRowCount1Based;

            requests.Add(new Request
            {
                UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
                {
                    Range = new DimensionRange { SheetId = sheetId, Dimension = "COLUMNS", StartIndex = 0, EndIndex = 1 },
                    Properties = new DimensionProperties { PixelSize = 260 },
                    Fields = "pixelSize"
                }
            });
            requests.Add(new Request
            {
                UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
                {
                    Range = new DimensionRange { SheetId = sheetId, Dimension = "COLUMNS", StartIndex = 1, EndIndex = 2 },
                    Properties = new DimensionProperties { PixelSize = 120 },
                    Fields = "pixelSize"
                }
            });
            requests.Add(new Request
            {
                UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
                {
                    Range = new DimensionRange { SheetId = sheetId, Dimension = "COLUMNS", StartIndex = 4, EndIndex = 5 },
                    Properties = new DimensionProperties { PixelSize = 120 },
                    Fields = "pixelSize"
                }
            });
            requests.Add(new Request
            {
                UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
                {
                    Range = new DimensionRange { SheetId = sheetId, Dimension = "COLUMNS", StartIndex = 5, EndIndex = 6 },
                    Properties = new DimensionProperties { PixelSize = 120 },
                    Fields = "pixelSize"
                }
            });

            requests.Add(new Request
            {
                UpdateSheetProperties = new UpdateSheetPropertiesRequest
                {
                    Properties = new SheetProperties
                    {
                        SheetId = sheetId,
                        GridProperties = new GridProperties { FrozenRowCount = 2 }
                    },
                    Fields = "gridProperties.frozenRowCount"
                }
            });

            // Kasaya satırı
            requests.Add(new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = 0,
                        EndRowIndex = 1,
                        StartColumnIndex = 0,
                        EndColumnIndex = 2
                    },
                    Cell = new CellData
                    {
                        UserEnteredFormat = new CellFormat
                        {
                            BackgroundColor = ColorFromHex(0xccfbf1),
                            VerticalAlignment = "MIDDLE",
                            TextFormat = new TextFormat { Bold = true, FontSize = 11 },
                            WrapStrategy = "WRAP"
                        }
                    },
                    Fields = "userEnteredFormat(backgroundColor,verticalAlignment,textFormat,wrapStrategy)"
                }
            });

            // Ürün başlığı
            requests.Add(new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = 1,
                        EndRowIndex = 2,
                        StartColumnIndex = 0,
                        EndColumnIndex = 2
                    },
                    Cell = new CellData
                    {
                        UserEnteredFormat = new CellFormat
                        {
                            BackgroundColor = ColorFromHex(0x0f766e),
                            HorizontalAlignment = "CENTER",
                            VerticalAlignment = "MIDDLE",
                            TextFormat = new TextFormat { Bold = true, FontSize = 11, ForegroundColor = ColorFromHex(0xffffff) },
                            WrapStrategy = "WRAP"
                        }
                    },
                    Fields = "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat,wrapStrategy)"
                }
            });

            if (e > 0)
            {
                requests.Add(new Request
                {
                    RepeatCell = new RepeatCellRequest
                    {
                        Range = new GridRange
                        {
                            SheetId = sheetId,
                            StartRowIndex = 2,
                            EndRowIndex = 2 + e,
                            StartColumnIndex = 0,
                            EndColumnIndex = 2
                        },
                        Cell = new CellData
                        {
                            UserEnteredFormat = new CellFormat
                            {
                                BackgroundColor = ColorFromHex(0xf8fafc),
                                VerticalAlignment = "MIDDLE",
                                WrapStrategy = "WRAP"
                            }
                        },
                        Fields = "userEnteredFormat(backgroundColor,verticalAlignment,wrapStrategy)"
                    }
                });
            }

            // Kasa
            requests.Add(new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = kasaRow0,
                        EndRowIndex = kasaRow0 + 1,
                        StartColumnIndex = 0,
                        EndColumnIndex = 2
                    },
                    Cell = new CellData
                    {
                        UserEnteredFormat = new CellFormat
                        {
                            BackgroundColor = ColorFromHex(0xe0f2fe),
                            VerticalAlignment = "MIDDLE",
                            TextFormat = new TextFormat { Bold = true, FontSize = 11 },
                            WrapStrategy = "WRAP"
                        }
                    },
                    Fields = "userEnteredFormat(backgroundColor,verticalAlignment,textFormat,wrapStrategy)"
                }
            });

            void StyleSummaryRow(int row0, uint bg)
            {
                requests.Add(new Request
                {
                    RepeatCell = new RepeatCellRequest
                    {
                        Range = new GridRange
                        {
                            SheetId = sheetId,
                            StartRowIndex = row0,
                            EndRowIndex = row0 + 1,
                            StartColumnIndex = 0,
                            EndColumnIndex = 2
                        },
                        Cell = new CellData
                        {
                            UserEnteredFormat = new CellFormat
                            {
                                BackgroundColor = ColorFromHex(bg),
                                VerticalAlignment = "MIDDLE",
                                TextFormat = new TextFormat { Bold = true, FontSize = 11 },
                                WrapStrategy = "WRAP"
                            }
                        },
                        Fields = "userEnteredFormat(backgroundColor,verticalAlignment,textFormat,wrapStrategy)"
                    }
                });
            }

            StyleSummaryRow(yuklenenRow0, 0xf1f5f9);
            StyleSummaryRow(devredecekRow0, 0xf1f5f9);
            StyleSummaryRow(devirRow0, 0xf1f5f9);
            StyleSummaryRow(toplamRow0, 0xfef3c7);
            StyleSummaryRow(devredenRow0, 0xfee2e2);
            StyleSummaryRow(satilanRow0, 0xffedd5);
            StyleSummaryRow(gunsonuRow0, 0xd9f99d);
            StyleSummaryRow(ciroRow0, 0xe9d5ff);

            if (mainEnd0 > 0)
            {
                requests.Add(new Request
                {
                    RepeatCell = new RepeatCellRequest
                    {
                        Range = new GridRange
                        {
                            SheetId = sheetId,
                            StartRowIndex = 0,
                            EndRowIndex = mainEnd0,
                            StartColumnIndex = 1,
                            EndColumnIndex = 2
                        },
                        Cell = new CellData
                        {
                            UserEnteredFormat = new CellFormat
                            {
                                HorizontalAlignment = "RIGHT",
                                NumberFormat = new NumberFormat { Type = "NUMBER", Pattern = "#,##0" }
                            }
                        },
                        Fields = "userEnteredFormat(horizontalAlignment,numberFormat)"
                    }
                });
            }

            var borderColor = ColorFromHex(0xcbd5e1);
            requests.Add(new Request
            {
                UpdateBorders = new UpdateBordersRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = 0,
                        EndRowIndex = mainEnd0,
                        StartColumnIndex = 0,
                        EndColumnIndex = 2
                    },
                    Top = new Border { Style = "SOLID", Width = 2, Color = borderColor },
                    Bottom = new Border { Style = "SOLID", Width = 2, Color = borderColor },
                    Left = new Border { Style = "SOLID", Width = 2, Color = borderColor },
                    Right = new Border { Style = "SOLID", Width = 2, Color = borderColor },
                    InnerHorizontal = new Border { Style = "SOLID", Width = 1, Color = borderColor },
                    InnerVertical = new Border { Style = "SOLID", Width = 1, Color = borderColor }
                }
            });

            // E:F özet aynası
            requests.Add(new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = 0,
                        EndRowIndex = MirrorRowDevirAlinan,
                        StartColumnIndex = 4,
                        EndColumnIndex = 6
                    },
                    Cell = new CellData
                    {
                        UserEnteredFormat = new CellFormat
                        {
                            BackgroundColor = ColorFromHex(0xf8fafc),
                            VerticalAlignment = "MIDDLE",
                            WrapStrategy = "WRAP"
                        }
                    },
                    Fields = "userEnteredFormat(backgroundColor,verticalAlignment,wrapStrategy)"
                }
            });

            requests.Add(new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = 0,
                        EndRowIndex = MirrorRowDevirAlinan,
                        StartColumnIndex = 4,
                        EndColumnIndex = 5
                    },
                    Cell = new CellData
                    {
                        UserEnteredFormat = new CellFormat
                        {
                            TextFormat = new TextFormat { Bold = true, FontSize = 10 },
                            HorizontalAlignment = "LEFT"
                        }
                    },
                    Fields = "userEnteredFormat(textFormat,horizontalAlignment)"
                }
            });

            requests.Add(new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = 0,
                        EndRowIndex = MirrorRowDevirAlinan,
                        StartColumnIndex = 5,
                        EndColumnIndex = 6
                    },
                    Cell = new CellData
                    {
                        UserEnteredFormat = new CellFormat
                        {
                            HorizontalAlignment = "RIGHT",
                            NumberFormat = new NumberFormat { Type = "NUMBER", Pattern = "#,##0" },
                            TextFormat = new TextFormat { Bold = true, FontSize = 11 }
                        }
                    },
                    Fields = "userEnteredFormat(horizontalAlignment,numberFormat,textFormat)"
                }
            });

            requests.Add(new Request
            {
                UpdateBorders = new UpdateBordersRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = 0,
                        EndRowIndex = MirrorRowDevirAlinan,
                        StartColumnIndex = 4,
                        EndColumnIndex = 6
                    },
                    Top = new Border { Style = "SOLID", Width = 2, Color = ColorFromHex(0x0f766e) },
                    Bottom = new Border { Style = "SOLID", Width = 2, Color = borderColor },
                    Left = new Border { Style = "SOLID", Width = 2, Color = borderColor },
                    Right = new Border { Style = "SOLID", Width = 2, Color = borderColor },
                    InnerHorizontal = new Border { Style = "SOLID", Width = 1, Color = borderColor },
                    InnerVertical = new Border { Style = "SOLID", Width = 1, Color = borderColor }
                }
            });

            return requests;
        }

        private static Color ColorFromHex(uint rgb)
        {
            var r = (rgb >> 16) & 0xff;
            var g = (rgb >> 8) & 0xff;
            var b = rgb & 0xff;
            return new Color
            {
                Red = r / 255f,
                Green = g / 255f,
                Blue = b / 255f
            };
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
