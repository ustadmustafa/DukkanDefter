namespace DukkanDefterOCR.Models
{
    public class OCRResultViewModel
    {
        public string Date { get; set; } = "";
        public string RawText { get; set; } = "";
        public string ImagePath { get; set; } = "";
        public List<OCRItem> Items { get; set; } = new();

        /// <summary>Metinde "ürün - tutar" biçiminde algılanan satır sayısı (toplam satırı hariç).</summary>
        public int DetectedDataRowCount { get; set; }
    }
}
