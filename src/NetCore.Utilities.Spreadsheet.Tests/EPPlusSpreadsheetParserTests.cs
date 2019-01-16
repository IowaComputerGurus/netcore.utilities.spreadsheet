using System;
using System.IO;
using Xunit;

namespace ICG.NetCore.Utilities.Spreadsheet.Tests
{
    public class EPPlusSpreadsheetParserTests
    {
        private readonly ISpreadsheetParser _spreadsheetParser;

        public EPPlusSpreadsheetParserTests()
        {
            _spreadsheetParser = new EPPlusSpreadsheetParser();
        }


        [Fact]
        public void ParseDocument_ShouldThrowArgumentExceptionWhenNoColumnIdentifiedForImport()
        {
            //Arrange

            //Act/Assert
            Assert.Throws<ArgumentException>(() =>
                _spreadsheetParser.ParseDocument<SampleExportRecord>(new MemoryStream()));
        }
    }
}