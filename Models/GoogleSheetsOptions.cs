namespace DukkanDefterOCR.Models
{
    public class GoogleSheetsOptions
    {
        public const string SectionName = "GoogleSheets";

        /// <summary>Servis hesabı JSON dosyasının tam yolu.</summary>
        public string CredentialsPath { get; set; } = "";

        /// <summary>Google Spreadsheet kimliği (URL içindeki ID).</summary>
        public string SpreadsheetId { get; set; } = "";
    }
}
