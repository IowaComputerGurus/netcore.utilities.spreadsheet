using System;
using System.Collections.Generic;
using System.IO;
using Bogus;
using FluentAssertions;
using Xunit;

namespace ICG.NetCore.Utilities.Spreadsheet.Tests;

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
        var configuration = new SpreadsheetConfiguration<SampleExportRecord> { WorksheetName = "    " };

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
        var configuration = new SpreadsheetConfiguration<SampleExportRecord> { WorksheetName = "TestSheet" };

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
            WorksheetName = "TestSheet",
            ExportData = new List<SampleExportRecord>(),
            RenderTitle = true
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

    private static Faker<TestExportRecord> GetTestExportRecordFaker() =>
        new Faker<TestExportRecord>()
            .UseSeed(1415)
            .RuleFor(i => i.DateTimeOffsetValue, f => f.Date.SoonOffset())
            .RuleFor(i => i.DateTimeValue, f => f.Date.Soon())
            .RuleFor(i => i.DecimalValue, f => f.Random.Decimal())
            .RuleFor(i => i.DoubleValue, f => f.Random.Double())
            .RuleFor(i => i.IntValue, f => f.Random.Int())
            .RuleFor(i => i.LongValue, f => f.Random.Long())
            .RuleFor(i => i.StringValue, f => f.Commerce.ProductName())
            .RuleFor(i => i.NullableDateTimeOffsetValue, f => f.Date.SoonOffset().OrNull(f))
            .RuleFor(i => i.NullableDateTimeValue, f => f.Date.Soon().OrNull(f))
            .RuleFor(i => i.NullableDecimalValue, f => f.Random.Decimal().OrNull(f))
            .RuleFor(i => i.NullableDoubleValue, f => f.Random.Double().OrNull(f))
            .RuleFor(i => i.NullableIntValue, f => f.Random.Int().OrNull(f))
            .RuleFor(i => i.NullableLongValue, f => f.Random.Long().OrNull(f))
            .RuleFor(i => i.NullableStringValue, f => f.Commerce.ProductName().OrNull(f))
            .RuleFor(i => i.DateOnly, f => f.Date.Soon())
            .RuleFor(i => i.Currency, f => decimal.Parse(f.Commerce.Price()))
            .RuleFor(i => i.Fixed0, f => f.Random.Decimal(0, 100))
            .RuleFor(i => i.Fixed1, f => f.Random.Decimal(0, 100))
            .RuleFor(i => i.Fixed2, f => f.Random.Decimal(0, 100))
            .RuleFor(i => i.Fixed3, f=> f.Random.Decimal(0, 100))
    ;

    private static Faker<DifferentTestExportRecord> GetDifferentTestExportRecordFaker() =>
        new Faker<DifferentTestExportRecord>()
            .UseSeed(9265)
            .RuleFor(i => i.Id, f => f.IndexFaker)
            .RuleFor(o => o.Company, f => f.Company.CompanyName())
            .RuleFor(o => o.Amount, f => f.Random.Double(0, 10))
            .RuleFor(o => o.Date, f => f.Date.Soon());

    [Fact]
    public void CreateMultiSheetSpreadsheet_With_A_Stream_Should_Work()
    {
        var testSheet1Data = GetTestExportRecordFaker().Generate(100);
        var testSheet2Data = GetDifferentTestExportRecordFaker().Generate(100);

        var config = new ISpreadsheetConfiguration[]
        {
                new SpreadsheetConfiguration<TestExportRecord>
                {
                    WorksheetName = "Sheet 1", ExportData = testSheet1Data
                },
                new SpreadsheetConfiguration<DifferentTestExportRecord>
                {
                    WorksheetName = "Sheet 2", ExportData = testSheet2Data
                }
        };
        using var ms = new MemoryStream();
        var result = _spreadsheetGenerator.CreateMultiSheetSpreadsheet(ms, config);

        result.Should().BeTrue();
        ms.Should().NotHaveLength(0);

        //var sheetPath = Path.Join(Path.GetTempPath(), $"CreateMultiSheetSpreadsheet_With_A_Stream_Should_Work_{DateTime.Now:yyyyMMddHHmmssfff}.xlsx");
        //File.WriteAllBytes(sheetPath, ms.ToArray());


        ms.Seek(0, SeekOrigin.Begin);
        ValidateExportedSheet(ms, 1, testSheet1Data);
        ms.Seek(0, SeekOrigin.Begin);
        ValidateExportedSheet(ms, 2, testSheet2Data);
    }

    [Fact]
    public void CreateMultiSheetSpreadsheet_Returning_Bytes_Should_Work()
    {
        var testSheet1Data = GetTestExportRecordFaker().Generate(100);
        var testSheet2Data = GetDifferentTestExportRecordFaker().Generate(100);

        var config = new MultisheetConfiguration()
            .WithSheet("Sheet 1", testSheet1Data, conf =>
            {
                conf.DocumentTitle = "An Amazing Title";
                conf.DocumentSubTitle = "An Even More Amazing Subtitle";
            })
            .WithSheet("Sheet 2", testSheet2Data);
        var result = _spreadsheetGenerator.CreateMultiSheetSpreadsheet(config);
        result.Should().NotBeNullOrEmpty();
        //var sheetPath = Path.Join(Path.GetTempPath(), $"CreateMultiSheetSpreadsheet_With_A_Stream_Should_Work_{DateTime.Now:yyyyMMddHHmmssfff}.xlsx");
        //File.WriteAllBytes(sheetPath, ms.ToArray());

        using var ms = new MemoryStream(result);
        ValidateExportedSheet(ms, 1, testSheet1Data);
        ms.Seek(0, SeekOrigin.Begin);
        ValidateExportedSheet(ms, 2, testSheet2Data);
    }

    [Fact]
    public void CreateSingleWorksheet_With_A_Stream_Should_Work()
    {
        using var ms = new MemoryStream();
        var testData = GetTestExportRecordFaker().Generate(100);

        var result = _spreadsheetGenerator.CreateSingleSheetSpreadsheet(ms, new SpreadsheetConfiguration<TestExportRecord>
        {
            WorksheetName = "Test Sheet",
            AutoSizeColumns = false,
            ExportData = testData
        });
        ms.Seek(0, SeekOrigin.Begin);
        ms.Should().NotHaveLength(0);
        ValidateExportedSheet(ms, 1, testData);
    }

    [Fact]
    public void CreateSingleWorksheet_Returning_Bytes_Should_Work()
    {
        var testData = GetTestExportRecordFaker().Generate(100);

        var result = _spreadsheetGenerator.CreateSingleSheetSpreadsheet(new SpreadsheetConfiguration<TestExportRecord>
        {
            WorksheetName = "Test Sheet",
            AutoSizeColumns = false,
            ExportData = testData
        });
        result.Should().NotBeNullOrEmpty();
        using var ms = new MemoryStream(result);
        /* These lines are write out the Excel spreadsheet that was generated for a manual check to ensure
         * that its actually displaying the values right, and we didn't offend its delicate sensibilities by
         * setting an attribute in the wrong place or something.
         */
        var sheetPath = Path.Join(Path.GetTempPath(), $"createsingleworksheet_should_work_{DateTime.Now:yyyyMMddHHmmssfff}.xlsx");
        File.WriteAllBytes(sheetPath, result);
        ValidateExportedSheet(ms, 1, testData);
    }
    
    [Fact]
    public void CreateSingleWorksheet_With_Formula()
    {
        using var ms = new MemoryStream();
        var testData = new List<SampleExportRecordWithFormula>()
        {
            new() { RecordTitle = "test record 1", RecordAmount = 10, RecordSize = 444 },
            new() { RecordTitle = "test record 2", RecordAmount = 20, RecordSize = 5 },
            new() { RecordTitle = "test record 3", RecordAmount = 30, RecordSize = 67 },
            new() { RecordTitle = "test record 4", RecordAmount = 3, RecordSize = 477 },
        };

        _spreadsheetGenerator.CreateSingleSheetSpreadsheet(ms, new SpreadsheetConfiguration<SampleExportRecordWithFormula>
        {
            WorksheetName = "Test Sheet",
            AutoSizeColumns = true,
            ExportData = testData
        });
        
        ms.Seek(0, SeekOrigin.Begin);
        ms.Should().NotHaveLength(0);
    }

    /*
     * This uses the OpenXmlSpreadsheetParser to attempt to pull the data out of the spreadsheet we just
     * exported. This ensures data was exported correctly, and has the added benefit of ensuring that
     * the parser can round trip the data written by the generator.
     */
    private static void ValidateExportedSheet<T>(Stream stream, int sheetNumber, IEnumerable<T> dataSet) where T : new()
    {
        var parser = new OpenXmlSpreadsheetParser();

        var parsed = parser.ParseDocument<T>(stream, sheetNumber, true);

        /*
         * Datetimes in Excel are fun. Especially if they get written out as OLE Automation values (Floating point days since 1900-1-1
         * with the fraction being how far through the day it is). Tell FluentAssertions we just care about second resolution. 
         */
        parsed.Should().BeEquivalentTo(dataSet, opt =>
                opt
                    .Using<DateTime>(ctx => ctx.Subject.Should().BeCloseTo(ctx.Expectation, TimeSpan.FromSeconds(1)))
                    .WhenTypeIs<DateTime>()
                    .Using<DateTimeOffset>(ctx => ctx.Subject.Should().BeCloseTo(ctx.Expectation, TimeSpan.FromSeconds(1)))
                    .WhenTypeIs<DateTimeOffset>()
                    .Using<DateTime?>(ctx =>
                    {
                        if (ctx.Expectation.HasValue)
                        {
                            ctx.Subject.Should().BeCloseTo(ctx.Expectation.Value, TimeSpan.FromSeconds(1));
                        }
                        else
                        {
                            ctx.Subject.Should().BeNull();
                        }
                    })
                    .WhenTypeIs<DateTime?>()
                    .Using<DateTimeOffset?>(ctx =>
                    {
                        if (ctx.Expectation.HasValue)
                        {
                            ctx.Subject.Should().BeCloseTo(ctx.Expectation.Value, TimeSpan.FromSeconds(1));
                        }
                        else
                        {
                            ctx.Subject.Should().BeNull();
                        }
                    })
                    .WhenTypeIs<DateTimeOffset?>()
                );
    }
}

