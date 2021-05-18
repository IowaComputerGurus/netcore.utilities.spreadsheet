using System;
using System.ComponentModel;
using ICG.NetCore.Utilities.Spreadsheet;

namespace NetCore.Utilities.SpreadsheetExample.Models
{
    public class SimpleExportData
    {
        public string Title { get; set; }

        [DisplayName("Due Date")]
        [SpreadsheetColumnFormat("D")]
        public DateTime DueDate { get; set; }

        [DisplayName("Total Cost")]
        [SpreadsheetColumnFormat("C")]
        public decimal TotalCost { get; set; }

        public string Notes { get; set; }
    }
}