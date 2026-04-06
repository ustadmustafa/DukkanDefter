using System.ComponentModel.DataAnnotations;

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

        [Required(ErrorMessage = "Devreden Akbil girin.")]
        [Display(Name = "Devreden Akbil")]
        public int? DevredenAkbil { get; set; }

        [Required(ErrorMessage = "Akbil girin.")]
        [Display(Name = "Akbil")]
        public int? Akbil { get; set; }

        [Display(Name = "Kasadan para çıkmadan yapılan akbil")]
        public bool KasadanParaCikmadan { get; set; }

        [Display(Name = "Kasadan para çıkmadan tutar")]
        public int? KasadanParaCikmadanTutar { get; set; }

        public List<ManualSheetRow> Rows { get; set; } = new();
    }
}
