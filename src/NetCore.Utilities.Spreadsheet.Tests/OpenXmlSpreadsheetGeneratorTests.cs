using System;
using System.Collections.Generic;
using Xunit;

namespace ICG.NetCore.Utilities.Spreadsheet.Tests
{
    // ReSharper disable once InconsistentNaming
    public class OpenXmlSpreadsheetGeneratorTests
    {
        private readonly ISpreadsheetGenerator _spreadsheetGenerator;

        public OpenXmlSpreadsheetGeneratorTests()
        {
            _spreadsheetGenerator = new OpenXmlSpreadsheetGenerator();
        }

        [Fact]
        public void CreateSingleSheetWorksheet_ShouldThrowArgumentException_WhenWorksheetNameIsMissing()
        {
            //Arrange
            var configuration = new SpreadsheetConfiguration<SampleExportRecord>();

            //Act/Assert
            //Assert that it throws the exception
            var result = Assert.Throws<ArgumentException>(() =>
                _spreadsheetGenerator.CreateSingleSheetSpreadsheet<SampleExportRecord>(configuration));
            //Assert for the proper property
            Assert.Equal("WorksheetName", result.ParamName);
        }

        [Fact]
        public void CreateSingleSheetWorksheet_ShouldThrowArgumentException_WhenWorksheetNameWhiteSpace()
        {
            //Arrange
            var configuration = new SpreadsheetConfiguration<SampleExportRecord>{WorksheetName = "    "};

            //Act/Assert
            //Assert that it throws the exception
            var result = Assert.Throws<ArgumentException>(() =>
                _spreadsheetGenerator.CreateSingleSheetSpreadsheet<SampleExportRecord>(configuration));
            //Assert for the proper property
            Assert.Equal("WorksheetName", result.ParamName);
        }

        [Fact]
        public void CreateSingleSheetWorksheet_ShouldThrowArgumentException_WhenExportDataIsNull()
        {
            //Arrange
            var configuration = new SpreadsheetConfiguration<SampleExportRecord>{WorksheetName = "TestSheet"};

            //Act/Assert
            //Assert that it throws the exception
            var result = Assert.Throws<ArgumentException>(() =>
                _spreadsheetGenerator.CreateSingleSheetSpreadsheet<SampleExportRecord>(configuration));
            //Assert for the proper property
            Assert.Equal("ExportData", result.ParamName);
        }

        [Fact]
        public void CreateSingleSheetWorksheet_ShouldThrowArgumentException_WhenRenderTitleIsTrueAndDocumentTitleIsNull()
        {
            //Arrange
            var configuration = new SpreadsheetConfiguration<SampleExportRecord>
            {
                WorksheetName = "TestSheet", ExportData = new List<SampleExportRecord>(), RenderTitle = true
            };

            //Act/Assert
            //Assert that it throws the exception
            var result = Assert.Throws<ArgumentException>(() =>
                _spreadsheetGenerator.CreateSingleSheetSpreadsheet<SampleExportRecord>(configuration));
            //Assert for the proper property
            Assert.Equal("DocumentTitle", result.ParamName);
        }

        [Fact]
        public void CreateSingleSheetWorksheet_ShouldThrowArgumentException_WhenRenderSubTitleIsTrueAndDocumentSubTitleIsNull()
        {
            //Arrange
            var configuration = new SpreadsheetConfiguration<SampleExportRecord>
            {
                WorksheetName = "TestSheet",
                ExportData = new List<SampleExportRecord>(),
                RenderSubTitle = true
            };

            //Act/Assert
            //Assert that it throws the exception
            var result = Assert.Throws<ArgumentException>(() =>
                _spreadsheetGenerator.CreateSingleSheetSpreadsheet<SampleExportRecord>(configuration));
            //Assert for the proper property
            Assert.Equal("DocumentSubTitle", result.ParamName);
        }

        [Fact]
        public void CreateMultiSheetSpreadsheet_ShouldThrowError_WhenProvidedSheetsIsNull()
        {
            //Arrange

            //Act
            var actualResult =
                Assert.Throws<ArgumentNullException>(() => _spreadsheetGenerator.CreateMultiSheetSpreadsheet(null));

            //Assert
            Assert.Equal("exportSheets", actualResult.ParamName);
        }
    }
}
