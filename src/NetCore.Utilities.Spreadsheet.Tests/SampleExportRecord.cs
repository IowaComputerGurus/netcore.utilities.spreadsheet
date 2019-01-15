using System.ComponentModel;

namespace ICG.NetCore.Utilities.Spreadsheet.Tests
{
    public class SampleExportRecord
    {
        [DisplayName("Title")]
        public string RecordTitle { get; set; }

        [DisplayName("Amount")]
        [SpreadsheetColumnFormat("C")]
        public decimal RecordAmount { get; set; }
    }
}