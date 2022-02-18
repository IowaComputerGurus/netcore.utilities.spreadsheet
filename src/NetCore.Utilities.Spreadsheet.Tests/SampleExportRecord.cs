using System;
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

    public class TestExportRecord
    {
        [SpreadsheetColumn(displayName:"Id", width: 15)]
        public int Id { get; set; }
        [SpreadsheetColumn(displayName: "Name")]
        public string Name { get; set; }
        [SpreadsheetColumn(displayName: "Date", format:"d")]
        public DateTime Date { get; set; }

        [SpreadsheetColumn(displayName: "Amount", format: "c")]
        public double Amount { get; set; }

    }
}