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
        private const string AkbilRowLabel = "Akbil";

        private readonly GoogleSheetsOptions _options;

        public GoogleSheetsService(IOptions<GoogleSheetsOptions> options)
        {
            _options = options.Value;
        }

        /// <param name="akbilHesapSatilan">Satılan Akbil formülünde kullanılacak toplam (Akbil + kasadan ek).</param>
        /// <param name="akbilTabloDeger">Sol tabloda gösterilecek, forma girilen Akbil tutarı.</param>
        public async Task SaveToSheetAsync(string sheetName, IList<OCRItem> items, int devredenAkbil, int akbilHesapSatilan, int akbilTabloDeger, CancellationToken ct = default)
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

            mainValues.Add(new List<object> { AkbilRowLabel, akbilTabloDeger });

            var mainRowCount = mainValues.Count;
            var footerStartRow = mainRowCount + 3;

            var previousDevreden = await TryReadDevredenAkbilFromPreviousDayAsync(service, sheetName, ct);
            var satilanAkbil = akbilHesapSatilan + (previousDevreden - devredenAkbil);

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

            spreadsheet = await service.Spreadsheets.Get(_options.SpreadsheetId).ExecuteAsync(ct);
            var sheet = spreadsheet.Sheets?.FirstOrDefault(s => s.Properties.Title == safeTitle);
            var sheetId = sheet?.Properties.SheetId;
            if (sheetId == null)
                return;

            var formatRequests = BuildLedgerFormatRequests(sheetId.Value, mainRowCount, footerStartRow);
            if (formatRequests.Count > 0)
            {
                await service.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest { Requests = formatRequests }, _options.SpreadsheetId).ExecuteAsync(ct);
            }
        }

        private static List<Request> BuildLedgerFormatRequests(int sheetId, int mainRowCount, int footerStartRow1Based)
        {
            var requests = new List<Request>();
            var footerRow0 = footerStartRow1Based - 1;

            // Sütun genişlikleri
            requests.Add(new Request
            {
                UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
                {
                    Range = new DimensionRange { SheetId = sheetId, Dimension = "COLUMNS", StartIndex = 0, EndIndex = 1 },
                    Properties = new DimensionProperties { PixelSize = 280 },
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
                    Range = new DimensionRange { SheetId = sheetId, Dimension = "COLUMNS", StartIndex = 2, EndIndex = 3 },
                    Properties = new DimensionProperties { PixelSize = 220 },
                    Fields = "pixelSize"
                }
            });
            requests.Add(new Request
            {
                UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
                {
                    Range = new DimensionRange { SheetId = sheetId, Dimension = "COLUMNS", StartIndex = 3, EndIndex = 4 },
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
                        GridProperties = new GridProperties { FrozenRowCount = 1 }
                    },
                    Fields = "gridProperties.frozenRowCount"
                }
            });

            // Başlık satırı
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
                            BackgroundColor = ColorFromHex(0x0f766e),
                            HorizontalAlignment = "CENTER",
                            VerticalAlignment = "MIDDLE",
                            TextFormat = new TextFormat
                            {
                                Bold = true,
                                FontSize = 11,
                                ForegroundColor = ColorFromHex(0xffffff)
                            },
                            WrapStrategy = "WRAP"
                        }
                    },
                    Fields = "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat,wrapStrategy)"
                }
            });

            // Ürün satırları (Toplam ve Akbil hariç)
            if (mainRowCount > 3)
            {
                requests.Add(new Request
                {
                    RepeatCell = new RepeatCellRequest
                    {
                        Range = new GridRange
                        {
                            SheetId = sheetId,
                            StartRowIndex = 1,
                            EndRowIndex = mainRowCount - 2,
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

            var toplamRowIndex = mainRowCount - 2;
            requests.Add(new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = toplamRowIndex,
                        EndRowIndex = toplamRowIndex + 1,
                        StartColumnIndex = 0,
                        EndColumnIndex = 2
                    },
                    Cell = new CellData
                    {
                        UserEnteredFormat = new CellFormat
                        {
                            BackgroundColor = ColorFromHex(0xfef3c7),
                            VerticalAlignment = "MIDDLE",
                            TextFormat = new TextFormat { Bold = true, FontSize = 11 },
                            WrapStrategy = "WRAP"
                        }
                    },
                    Fields = "userEnteredFormat(backgroundColor,verticalAlignment,textFormat,wrapStrategy)"
                }
            });

            var akbilRowIndex = mainRowCount - 1;
            requests.Add(new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = akbilRowIndex,
                        EndRowIndex = akbilRowIndex + 1,
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

            // Tutar sütunu sayı formatı (sol tablo)
            if (mainRowCount > 1)
            {
                requests.Add(new Request
                {
                    RepeatCell = new RepeatCellRequest
                    {
                        Range = new GridRange
                        {
                            SheetId = sheetId,
                            StartRowIndex = 1,
                            EndRowIndex = mainRowCount,
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

            // Sol tablo dış çerçeve + iç yatay çizgiler
            var borderColor = ColorFromHex(0xcbd5e1);
            requests.Add(new Request
            {
                UpdateBorders = new UpdateBordersRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = 0,
                        EndRowIndex = mainRowCount,
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

            // Akbil bloğu (C:D)
            requests.Add(new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = footerRow0,
                        EndRowIndex = footerRow0 + 2,
                        StartColumnIndex = 2,
                        EndColumnIndex = 4
                    },
                    Cell = new CellData
                    {
                        UserEnteredFormat = new CellFormat
                        {
                            BackgroundColor = ColorFromHex(0xf1f5f9),
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
                        StartRowIndex = footerRow0,
                        EndRowIndex = footerRow0 + 2,
                        StartColumnIndex = 2,
                        EndColumnIndex = 3
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
                        StartRowIndex = footerRow0,
                        EndRowIndex = footerRow0 + 2,
                        StartColumnIndex = 3,
                        EndColumnIndex = 4
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
                        StartRowIndex = footerRow0,
                        EndRowIndex = footerRow0 + 2,
                        StartColumnIndex = 2,
                        EndColumnIndex = 4
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

        /// <summary>RGB hex 0xRRGGBB → Google API Color (sRGB 0–1).</summary>
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
