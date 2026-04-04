namespace DukkanDefterOCR.Models
{
    public class OCRItem
    {
        public string Kalem { get; set; } = "";
        public string HamTutar { get; set; } = "";
        public int Toplam { get; set; }
        public string Not { get; set; } = "";

        /// <summary>OCR satırı değil; tablonun son satırındaki genel toplam.</summary>
        public bool IsGrandTotalRow { get; set; }
    }
}
