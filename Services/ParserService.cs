using System.Text;
using System.Text.RegularExpressions;
using DukkanDefterOCR.Models;

namespace DukkanDefterOCR.Services
{
    /// <summary>
    /// Defter formatı: "Ürün - tutar" veya "Ürün - 1000 + 2000"; en sonda TOPLAM satırı.
    /// </summary>
    public class ParserService
    {
        private static readonly Regex DateRx = new(@"(?<!\d)(\d{2}[./]\d{2}[./]\d{2,4})(?!\d)", RegexOptions.Compiled);
        private static readonly Regex LineItemRx = new(@"^(.+?)\s*[-–—]\s*(.+)$", RegexOptions.Compiled);
        private static readonly Regex OnlyBigNumberRx = new(@"^\s*(\d{4,})\s*$", RegexOptions.Compiled);

        public OCRResultViewModel Parse(string text)
        {
            var normalized = NormalizeText(text);
            var result = new OCRResultViewModel { RawText = text };

            var lines = SplitPhysicalLines(normalized);
            var structureRowCount = lines.Count(IsStructureRowLine);

            var dataRows = new List<OCRItem>();
            string? footerTotalDigits = null;

            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                if (string.IsNullOrEmpty(result.Date))
                {
                    var dm = DateRx.Match(line);
                    if (dm.Success)
                    {
                        result.Date = dm.Value.Replace("/", ".");
                        if (line.Trim().Length <= 14 && !line.Contains('-', StringComparison.Ordinal))
                            continue;
                    }
                }

                var onlyNum = OnlyBigNumberRx.Match(line);
                if (onlyNum.Success && !line.Contains('-', StringComparison.Ordinal))
                {
                    footerTotalDigits = onlyNum.Groups[1].Value;
                    continue;
                }

                var m = LineItemRx.Match(line);
                if (!m.Success)
                    continue;

                var kalem = CleanLabel(m.Groups[1].Value);
                var valuePart = m.Groups[2].Value.Trim();
                if (string.IsNullOrEmpty(kalem))
                    continue;

                var note = "";
                var noteMatch = Regex.Match(valuePart, @"\((.*?)\)");
                if (noteMatch.Success)
                {
                    note = noteMatch.Groups[1].Value.Trim();
                    valuePart = Regex.Replace(valuePart, @"\(.*?\)", "", RegexOptions.Singleline).Trim();
                }

                var toplam = SumMoneyTokens(valuePart);

                dataRows.Add(new OCRItem
                {
                    Kalem = kalem,
                    HamTutar = valuePart,
                    Toplam = toplam,
                    Not = note,
                    IsGrandTotalRow = false
                });
            }

            if (string.IsNullOrWhiteSpace(result.Date))
                result.Date = "Tarihsiz";

            result.DetectedDataRowCount = Math.Max(structureRowCount, dataRows.Count);

            while (dataRows.Count < structureRowCount)
            {
                dataRows.Add(new OCRItem
                {
                    Kalem = "",
                    HamTutar = "",
                    Toplam = 0,
                    Not = "",
                    IsGrandTotalRow = false
                });
            }

            var dataSum = dataRows.Sum(i => i.Toplam);
            dataRows.Add(new OCRItem
            {
                Kalem = "TOPLAM",
                HamTutar = footerTotalDigits != null ? $"Sayfa altı okunan: {footerTotalDigits}" : "",
                Toplam = dataSum,
                Not = "",
                IsGrandTotalRow = true
            });

            result.Items = dataRows;
            return result;
        }

        private static string NormalizeText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return "";

            var sb = new StringBuilder(text.Length);
            foreach (var c in text)
            {
                if (c == '\r') continue;
                if (c == '—' || c == '–') sb.Append('-');
                else sb.Append(c);
            }
            return sb.ToString();
        }

        private static List<string> SplitPhysicalLines(string text)
        {
            return text
                .Split('\n', StringSplitOptions.None)
                .Select(l => l.Trim())
                .Where(l => !string.IsNullOrWhiteSpace(l))
                .ToList();
        }

        private static bool IsStructureRowLine(string line)
        {
            if (OnlyBigNumberRx.IsMatch(line) && !line.Contains('-', StringComparison.Ordinal))
                return false;
            if (DateRx.IsMatch(line) && !line.Contains('-', StringComparison.Ordinal) && line.Length < 16)
                return false;
            return LineItemRx.IsMatch(line);
        }

        private static string CleanLabel(string raw)
        {
            var s = raw.Trim();
            s = Regex.Replace(s, @"^\W+", "");
            s = Regex.Replace(s, @"\W+$", "");
            return s.Trim();
        }

        private static int SumMoneyTokens(string valuePart)
        {
            if (string.IsNullOrWhiteSpace(valuePart))
                return 0;

            var cleaned = valuePart
                .Replace(" + ", "+", StringComparison.Ordinal)
                .Replace(" +", "+", StringComparison.Ordinal)
                .Replace("+ ", "+", StringComparison.Ordinal);

            cleaned = Regex.Replace(cleaned, @"[^\d+]", "+");
            cleaned = Regex.Replace(cleaned, @"\+{2,}", "+");
            cleaned = cleaned.Trim('+');

            if (string.IsNullOrEmpty(cleaned))
                return 0;

            var sum = 0;
            foreach (var part in cleaned.Split('+', StringSplitOptions.RemoveEmptyEntries))
            {
                if (int.TryParse(part.Trim(), out var n))
                    sum += n;
            }

            return sum;
        }
    }
}
