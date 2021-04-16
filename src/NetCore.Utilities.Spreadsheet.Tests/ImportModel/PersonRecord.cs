namespace ICG.NetCore.Utilities.Spreadsheet.Tests.ImportModel
{
    public class PersonRecord
    {
        [SpreadsheetImportColumn(1)]
        public string Name { get; set; }

        [SpreadsheetImportColumn(2)]
        public int Age { get; set; }
    }
}