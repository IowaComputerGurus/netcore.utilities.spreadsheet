using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using FontSize = DocumentFormat.OpenXml.Spreadsheet.FontSize;

namespace ICG.NetCore.Utilities.Spreadsheet;
#nullable enable
/// <summary>
/// A concrete implementation of <see cref="ISpreadsheetGenerator"/> using the OpenXML project
/// </summary>
public class OpenXmlSpreadsheetGenerator : ISpreadsheetGenerator
{

    /// <inheritdoc/>
    public byte[] CreateSingleSheetSpreadsheet<T>(SpreadsheetConfiguration<T> exportConfiguration) where T : class
    {
        using var ms = new MemoryStream();
        var result = CreateSingleSheetSpreadsheet(ms, exportConfiguration);
        if (!result) return Array.Empty<byte>();
        ms.Seek(0, SeekOrigin.Begin);
        return ms.ToArray();
    }

    /// <inheritdoc/>
    public bool CreateSingleSheetSpreadsheet<T>(Stream output, SpreadsheetConfiguration<T> exportConfiguration) where T : class
    {
        //Validate input
        if (string.IsNullOrWhiteSpace(exportConfiguration.WorksheetName))
            throw new ArgumentException("Worksheet name must be supplied", nameof(exportConfiguration.WorksheetName));

        if (exportConfiguration.ExportData == null)
            throw new ArgumentException("Export data must be specified", nameof(exportConfiguration.ExportData));

        if (exportConfiguration.RenderTitle && string.IsNullOrEmpty(exportConfiguration.DocumentTitle))
            throw new ArgumentException("Document Title is required when 'Render Title' is true",
                nameof(exportConfiguration.DocumentTitle));

        if (exportConfiguration.RenderSubTitle && string.IsNullOrEmpty(exportConfiguration.DocumentSubTitle))
            throw new ArgumentException("Document Sub Title is required when 'Render Sub Title' is true",
                nameof(exportConfiguration.DocumentSubTitle));

        //Create the document & overall workbook
        var spreadsheetDocument = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook);

        var workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        Debug.Assert(spreadsheetDocument.WorkbookPart != null, "spreadsheetDocument.WorkbookPart != null. Something must be wrong with the universe.");

        //Setup our styles
        var stylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateStylesheet();
        stylesPart.Stylesheet.Save();

        var data = CreateExportSheet(exportConfiguration, out var columns);

        //Add a worksheet to our document
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet();
        worksheetPart.Worksheet.Append(columns);
        worksheetPart.Worksheet.Append(data);

        //Add the sheet to the workbook
        var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
        var sheet = new Sheet
        {
            Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = exportConfiguration.WorksheetName
        };
        sheets.Append(sheet);

        workbookPart.Workbook.Save();
        spreadsheetDocument.Close();
        return true;

    }

    /// <inheritdoc />
    public byte[] CreateMultiSheetSpreadsheet(MultisheetConfiguration configuration) 
        => CreateMultiSheetSpreadsheet((IEnumerable<ISpreadsheetConfiguration>)configuration);

    /// <inheritdoc />
    public bool CreateMultiSheetSpreadsheet(Stream output, MultisheetConfiguration configuration)
        => CreateMultiSheetSpreadsheet(output, (IEnumerable<ISpreadsheetConfiguration>)configuration);

    /// <inheritdoc />
    public byte[] CreateMultiSheetSpreadsheet(IEnumerable<ISpreadsheetConfiguration> exportSheets)
    {
        using var ms = new MemoryStream();
        var result = CreateMultiSheetSpreadsheet(ms, exportSheets);
        if (!result) return Array.Empty<byte>();
        ms.Seek(0, SeekOrigin.Begin);
        return ms.ToArray();
    }

    /// <inheritdoc />
    public bool CreateMultiSheetSpreadsheet(Stream output, IEnumerable<ISpreadsheetConfiguration> exportSheets)
    {
        //Validate inputs
        if (exportSheets == null)
            throw new ArgumentNullException(nameof(exportSheets));

        //Create the document & overall workbook
        var spreadsheetDocument = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook);
        var workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

        //Setup our styles
        var stylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateStylesheet();
        stylesPart.Stylesheet.Save();

        //Loop through all of the sheets
        var sheetId = 1u;
        foreach (var item in exportSheets)
        {
            var data = CreateExportSheet(item, out var columns);

            //Add a worksheet to our document
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();
            worksheetPart.Worksheet.Append(columns);
            worksheetPart.Worksheet.Append(data);

            //Add the sheet to the workbook
            var sheet = new Sheet
            {
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetId,
                Name = item.WorksheetName
            };
            sheets.Append(sheet);
            sheetId++;
        }
        //Save off the file
        workbookPart.Workbook.Save();
        spreadsheetDocument.Close();

        return true;
    }

    private static bool IsOfType<T>(Type t)
    {
        var typeToCheck = typeof(T);

        var unwrapped = Nullable.GetUnderlyingType(t);
        if (unwrapped == null)
        {
            return t == typeToCheck;
        }

        return unwrapped == typeToCheck;
    }

    private static Cell CellFromValue(Type t, object? itemValue) => t switch
    {
        _ when itemValue == null => new Cell { CellValue = new CellValue() },
        _ when IsOfType<int>(t) => new Cell { CellValue = new CellValue((int)itemValue), DataType = CellValues.Number },
        _ when IsOfType<decimal>(t) => new Cell { CellValue = new CellValue((decimal)itemValue), DataType = CellValues.Number },
        _ when IsOfType<double>(t) => new Cell { CellValue = new CellValue((double)itemValue), DataType = CellValues.Number },
        _ when IsOfType<long>(t) => new Cell { CellValue = new CellValue(Convert.ToDecimal((long)itemValue)), DataType = CellValues.Number }, //There is no constructor for longs
        _ when IsOfType<float>(t) => new Cell { CellValue = new CellValue((float)itemValue), DataType = CellValues.Number },
        _ when IsOfType<DateTime>(t) => new Cell { CellValue = new CellValue(((DateTime)itemValue)), DataType = CellValues.Date, StyleIndex = (int)FontStyleIndex.DateTime},
        _ when IsOfType<DateTimeOffset>(t) => new Cell { CellValue = new CellValue((DateTimeOffset)itemValue), DataType = CellValues.String},
        _ => new Cell { CellValue = new CellValue(itemValue.ToString() ?? ""), DataType = CellValues.String },
    };

    private sealed record OutputPropMap(Column Column, List<Cell> Cells);

    private static SheetData CreateExportSheet(ISpreadsheetConfiguration exportConfiguration, out Columns columns)
    {
        //Build out our sheet information
        var data = new SheetData();
        UInt32 currentRow = 1;

        columns = new Columns();

        var outputMap = new Dictionary<PropDetail, OutputPropMap>();
        if (exportConfiguration.RenderTitle)
        {
            var row = new Row { RowIndex = 1 };
            var headerCell = new Cell
            {
                CellReference = $"A{currentRow}",
                CellValue = new CellValue(exportConfiguration.DocumentTitle),
                DataType = CellValues.String,
                StyleIndex = (int)FontStyleIndex.Header
            };
            row.Append(headerCell);
            data.Append(row);
            //Increment row
            currentRow++;
        }

        if (exportConfiguration.RenderSubTitle)
        {
            var row = new Row { RowIndex = currentRow };
            var headerCell = new Cell
            {
                CellReference = $"A{currentRow}",
                CellValue = new CellValue(exportConfiguration.DocumentSubTitle),
                DataType = CellValues.String,
                StyleIndex = (int)FontStyleIndex.SubHeader
            };
            row.Append(headerCell);
            data.Append(row);
            //Increment row
            currentRow++;
        }

        //Run data headers
        var headerProperties = TypeDiscoverer.GetProps(exportConfiguration.DataType);
        var headerRow = new Row { RowIndex = currentRow };
        foreach (var prop in headerProperties)
        {
            var column = new Column()
            {
                Min = (uint)prop.Order,
                Max = (uint)prop.Order,
                BestFit = true,
                Width = prop.Width > 0 ? prop.Width : 10
            };
            columns.Append(column);
            outputMap[prop] = new OutputPropMap(column, new List<Cell>());

            var headerCell = new Cell
            {
                CellValue = new CellValue(prop.DisplayName),
                DataType = CellValues.String,
                StyleIndex = (int)FontStyleIndex.DataHeader
            };
            headerRow.Append(headerCell);
            outputMap[prop].Cells.Add(headerCell);
        }

        data.Append(headerRow);
        currentRow++;

        //Run the data
        foreach (var item in exportConfiguration.ExportData)
        {
            var dataRow = new Row { RowIndex = currentRow };
            foreach (var prop in headerProperties)
            {
                var itemValue = prop.Descriptor.GetValue(item);
                var dataCell = CellFromValue(prop.Descriptor.PropertyType, itemValue);

                dataCell.StyleIndex = prop.Format switch
                {
                    ColumnFormats.Currency => (int)FontStyleIndex.NormalCurrency,
                    ColumnFormats.Date => (int)FontStyleIndex.NormalDate,
                    ColumnFormats.Fixed0 => (int)FontStyleIndex.Fixed0,
                    ColumnFormats.Fixed1 => (int)FontStyleIndex.Fixed1,
                    ColumnFormats.Fixed2 => (int)FontStyleIndex.Fixed2,
                    _ => dataCell.StyleIndex
                };
                outputMap[prop].Cells.Add(dataCell);
                dataRow.Append(dataCell);
            }

            data.Append(dataRow);
            currentRow++;
        }
        
        if (exportConfiguration.AutoSizeColumns)
        {
            CalculateSizes(outputMap.Values.ToList());
        }
        return data;
    }

    /// <summary>
    /// Internal helper for tracking font-style indexes
    /// </summary>
    private enum FontStyleIndex
    {
        /// <summary>
        /// Default cell text
        /// </summary>
        Default = 0,
        /// <summary>
        /// Document header formatting
        /// </summary>
        Header = 1,
        /// <summary>
        /// Document sub-header formatting
        /// </summary>
        SubHeader = 2,
        /// <summary>
        /// Headers for actual data.
        /// </summary>
        DataHeader = 3,
        /// <summary>
        /// Normal font formatted for currency
        /// </summary>
        NormalCurrency = 4,
        /// <summary>
        /// Normal font formatted for date
        /// </summary>
        NormalDate = 5,
        DateTime = 6,
        Fixed0 = 7,
        Fixed1=8,
        Fixed2=9,
    }

    /// <summary>
    /// Creates the needed stylesheet to support our styles for documents
    /// </summary>
    /// <returns></returns>
    private static Stylesheet CreateStylesheet()
    {
        var styles = new Stylesheet();
        // blank font list
        styles.Fonts = new Fonts { Count = 4 };
        styles.Fonts.AppendChild(new Font()
        {
            FontSize = new FontSize { Val = 11 }
        });
        styles.Fonts.AppendChild(new Font()
        {
            Bold = new Bold(),
            FontSize = new FontSize() { Val = 14 }
        });
        styles.Fonts.AppendChild(new Font()
        {
            Bold = new Bold(),
            FontSize = new FontSize() { Val = 12 }
        });
        styles.Fonts.AppendChild(new Font()
        {
            Bold = new Bold()
        });

        // create fills
        var solidRed = new PatternFill
        {
            PatternType = PatternValues.Solid,
            ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFF0000") },
            BackgroundColor = new BackgroundColor { Indexed = 64 }
        };
        styles.Fills = new Fills(
            new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } },
            new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } },
            new Fill { PatternFill = solidRed }
        );

        // blank border list
        styles.Borders = new Borders(new Border());

        // blank cell format list
        styles.CellStyleFormats = new CellStyleFormats(new CellFormat());

        styles.NumberingFormats = new NumberingFormats(
            new NumberingFormat { NumberFormatId = 164, FormatCode = "\"$\"#,##0.00" },
            new NumberingFormat { NumberFormatId = 300, FormatCode = "mm/dd/yyyy" },
            new NumberingFormat { NumberFormatId = 301, FormatCode = "0" },
            new NumberingFormat { NumberFormatId = 302, FormatCode = "0.0" },
            new NumberingFormat { NumberFormatId = 303, FormatCode = "0.00" }
        );

        styles.CellFormats = new CellFormats(
            new CellFormat(), //Default
            new CellFormat { FormatId = 0, FontId = 1, BorderId = 0, FillId = 0 },
            new CellFormat { FormatId = 0, FontId = 2, BorderId = 0, FillId = 0 },
            new CellFormat { FormatId = 0, FontId = 3, BorderId = 0, FillId = 0 },
            new CellFormat
            {
                FormatId = 0,
                FontId = 0,
                BorderId = 0,
                FillId = 0,
                NumberFormatId = 164,
                ApplyNumberFormat = true
            },
            new CellFormat
            {
                FormatId = 0,
                FontId = 0,
                BorderId = 0,
                FillId = 0,
                NumberFormatId = 300,
                ApplyNumberFormat = true
            },
            new CellFormat
            {
                NumberFormatId = 22,
                ApplyNumberFormat = true
            },
            new CellFormat
            {
                FormatId = 0,
                FontId = 0,
                BorderId = 0,
                FillId = 0,
                NumberFormatId = 301,
                ApplyNumberFormat = true
            },
            new CellFormat
            {
                FormatId = 0,
                FontId = 0,
                BorderId = 0,
                FillId = 0,
                NumberFormatId = 302,
                ApplyNumberFormat = true
            }, new CellFormat
            {
                FormatId = 0,
                FontId = 0,
                BorderId = 0,
                FillId = 0,
                NumberFormatId = 303,
                ApplyNumberFormat = true
            }
        );
        return styles;
    }


    private static void CalculateSizes(IList<OutputPropMap> propMap)
    {
        //Adapted from - https://stackoverflow.com/questions/18268620/openxml-auto-size-column-width-in-excel
        //This is an approximation of the size needed for the largest single character in Calibri 
        double maxWidth = 7;
        foreach (var (col, cells) in propMap)
        {
            var rawWidth = GetMaxCharacterWidth(cells);
            //width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
            var width = Math.Truncate((rawWidth * maxWidth + 5) / maxWidth * 256) / 256;
            col.CustomWidth = true;
            col.Width = width;
        }
    }

    private static int GetMaxCharacterWidth(IList<Cell> cells)
    {
        //iterate over all cells getting a max char value for each column
        var maxWidth = 0;

        //TODO: Be smarter about this for our set styles
        var numberStyles = new UInt32[] { 7, 8 }; //styles that will add extra chars
        var boldStyles = new UInt32[] { 1, 2, 3, 4, 6, 7, 8 }; //styles that will bold

        //using cell index as my column
        foreach (var cell in cells)
        {
            var cellValue = cell.CellValue == null ? string.Empty : cell.CellValue.InnerText;
            var cellTextLength = cellValue.Length;

            if (cell.StyleIndex?.HasValue ?? false)
            {
                if (numberStyles.Contains(cell.StyleIndex))
                {
                    int thousandCount = (int)Math.Truncate((double)cellTextLength / 4);

                    //add 3 for '.00' 
                    cellTextLength += (3 + thousandCount);
                }

                if (boldStyles.Contains(cell.StyleIndex))
                {
                    //add an extra char for bold - not 100% 
                    cellTextLength += 1;
                }
            }

            if (cellTextLength > maxWidth)
            {
                maxWidth = cellTextLength;
            }
        }

        return maxWidth;
    }
}