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

        [Required(ErrorMessage = "Kasaya girilen para girin.")]
        [Display(Name = "Kasaya Girilen Para")]
        public int? KasayaGirilenPara { get; set; }

        [Required(ErrorMessage = "Kasa tutarını girin.")]
        [Display(Name = "Kasa")]
        public int? Kasa { get; set; }

        [Required(ErrorMessage = "Yüklenen akbil girin.")]
        [Display(Name = "Yüklenen Akbil")]
        public int? YuklenenAkbil { get; set; }

        [Required(ErrorMessage = "Devredecek akbil girin.")]
        [Display(Name = "Devredecek Akbil")]
        public int? DevredecekAkbil { get; set; }

        /// <summary>Önceki günün sekmesinden okunan; yalnızca gösterim (kayıtta sunucu yeniden okur).</summary>
        [Display(Name = "Devir Alınan Akbil")]
        public int DevirAlinanAkbilPreview { get; set; }

        /// <summary>Önceki günün Kasa değeri; önizleme ve Ciro hesabı için.</summary>
        [Display(Name = "Dünkü Kasa (önceki gün)")]
        public int DunKasaFromPreviousDay { get; set; }

        public List<ManualSheetRow> Rows { get; set; } = new();
    }
}
