using System;
using System.ComponentModel;

namespace ICG.NetCore.Utilities.Spreadsheet.Tests;

#nullable enable
public class SampleExportRecord
{
    [DisplayName("Title")]
    public string? RecordTitle { get; set; }

    [DisplayName("Amount")]
    [SpreadsheetColumn("Amount", format:"C")]
    public decimal RecordAmount { get; set; }
}

public class TestExportRecord
{
    [SpreadsheetColumn(displayName: "IntValue")]
    [SpreadsheetImportColumn(1)]
    public int IntValue { get; set; }

    [SpreadsheetColumn(displayName: "NullableIntValue")]
    [SpreadsheetImportColumn(2)]
    public int? NullableIntValue { get; set; }

    [SpreadsheetColumn(displayName: "DecimalValue")]
    [SpreadsheetImportColumn(3)]
    public decimal DecimalValue { get; set; }

    [SpreadsheetColumn(displayName: "NullableDecimalValue")]
    [SpreadsheetImportColumn(4)]
    public decimal? NullableDecimalValue { get; set; }

    [SpreadsheetColumn(displayName: "DoubleValue")]
    [SpreadsheetImportColumn(5)]
    public double DoubleValue { get; set; }

    [SpreadsheetColumn(displayName: "NullableDoubleValue")]
    [SpreadsheetImportColumn(6)]
    public double? NullableDoubleValue { get; set; }

    [SpreadsheetColumn(displayName: "LongValue")]
    [SpreadsheetImportColumn(7)]
    public long LongValue { get; set; }

    [SpreadsheetColumn(displayName: "NullableLongValue")]
    [SpreadsheetImportColumn(8)]
    public long? NullableLongValue { get; set; }

    [SpreadsheetColumn(displayName: "DateTimeValue", format: "d")]
    [SpreadsheetImportColumn(9)]
    public DateTime DateTimeValue { get; set; }

    [SpreadsheetColumn(displayName: "NullableDateTimeValue")]
    [SpreadsheetImportColumn(10)]
    public DateTime? NullableDateTimeValue { get; set; }

    [SpreadsheetColumn(displayName: "DateTimeOffsetValue")]
    [SpreadsheetImportColumn(11)]
    public DateTimeOffset DateTimeOffsetValue { get; set; }

    [SpreadsheetColumn(displayName: "NullableDateTimeOffsetValue")]
    [SpreadsheetImportColumn(12)]
    public DateTimeOffset? NullableDateTimeOffsetValue { get; set; }

    [SpreadsheetColumn(displayName: "StringValue")]
    [SpreadsheetImportColumn(13)]
    public string StringValue { get; set; } = "";

    [SpreadsheetColumn(displayName: "NullableStringValue")]
    [SpreadsheetImportColumn(14)]
    public string? NullableStringValue { get; set; }
}

public class DifferentTestExportRecord
{
    [SpreadsheetColumn(displayName: "Id", width: 15)]
    [SpreadsheetImportColumn(1)]
    public int Id { get; set; }

    [SpreadsheetColumn(displayName: "Company")]
    [SpreadsheetImportColumn(2)]
    public string Company { get; set; } = "";

    [SpreadsheetColumn(displayName: "Date", format: "d")]
    [SpreadsheetImportColumn(3)]
    public DateTime Date { get; set; }

    [SpreadsheetColumn(displayName: "Amount", format: "c")]
    [SpreadsheetImportColumn(4)]
    public double Amount { get; set; }

}