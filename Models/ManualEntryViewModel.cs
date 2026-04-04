namespace DukkanDefterOCR.Models
{
    public class ManualSheetRow
    {
        public string Urun { get; set; } = "";
        public int Tutar { get; set; }
    }

    public class ManualEntryViewModel
    {
        /// <summary>Google Sheet sekmesi adı (genelde tarih).</summary>
        public string SheetDate { get; set; } = "";

        public List<ManualSheetRow> Rows { get; set; } = new();
    }
}
