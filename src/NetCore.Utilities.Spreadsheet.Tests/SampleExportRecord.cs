using System;
using System.ComponentModel;

namespace ICG.NetCore.Utilities.Spreadsheet.Tests;

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
    [SpreadsheetColumn("Id", 15)]
    public int Id { get; set; }

    [SpreadsheetColumn("Name")]
    public string Name { get; set; }

    [SpreadsheetColumn("Date", format: "d")]
    public DateTime Date { get; set; }

    [SpreadsheetColumn("Amount", format: "c")]
    public double Amount { get; set; }
}