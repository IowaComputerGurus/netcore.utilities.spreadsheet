using System;
using ICG.NetCore.Utilities.Spreadsheet;

namespace NetCore.Utilities.SpreadsheetExample.Models
{
    public class SimpleExportData
    {
        public string Title { get; set; }

        [SpreadsheetColumn("Due Date", format:"D")]
        public DateTime DueDate { get; set; }
        
        [SpreadsheetColumn("Total Cost", format:"C", formula: "SUM")]
        public decimal TotalCost { get; set; }

        [SpreadsheetColumn("Testing Numbers", format:"F3")]
        public decimal TestingNumbers { get; set; }

        public string Notes { get; set; }
    }
}