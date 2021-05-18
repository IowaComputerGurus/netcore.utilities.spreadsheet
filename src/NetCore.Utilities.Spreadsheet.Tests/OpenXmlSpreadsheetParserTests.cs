using System;
using System.IO;
using System.Linq;
using ICG.NetCore.Utilities.Spreadsheet.Tests.ImportModel;
using Xunit;

namespace ICG.NetCore.Utilities.Spreadsheet.Tests
{
    public class OpenXmlSpreadsheetParserTests
    {
        private readonly ISpreadsheetParser _spreadsheetParser;

        public OpenXmlSpreadsheetParserTests()
        {
            _spreadsheetParser = new OpenXmlSpreadsheetParser();
        }


        [Fact]
        public void ParseDocument_ShouldThrowArgumentExceptionWhenNoColumnIdentifiedForImport()
        {
            //Arrange

            //Act/Assert
            Assert.Throws<ArgumentException>(() =>
                _spreadsheetParser.ParseDocument<SampleExportRecord>(new MemoryStream()));
        }

        [Fact]
        public void ParseDocument_ShouldReturnProperData()
        {
            //Arrange
            var filePath = "../../../SampleFiles/ImportSample.xlsx";
            var expectedCount = 2;

            //Act
            var result = _spreadsheetParser.ParseDocument<PersonRecord>(File.OpenRead(filePath), 1, true);

            //Assert
            Assert.Equal(expectedCount, result.Count);
            var firstRecord = result.First();
            Assert.Equal("John Smith", firstRecord.Name);
            Assert.Equal(55, firstRecord.Age);
        }
    }
}