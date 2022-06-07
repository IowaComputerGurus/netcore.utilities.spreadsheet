using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ICG.NetCore.Utilities.Spreadsheet;
#nullable enable

/// <inheritdoc />
public class OpenXmlSpreadsheetParser : ISpreadsheetParser
{
    /// <inheritdoc />
    public List<T> ParseDocument<T>(Stream fileStream) where T : new()
    {
        return ParseDocument<T>(fileStream, 1);
    }

    /// <inheritdoc />
    public List<T> ParseDocument<T>(Stream fileStream, int worksheetNumber) where T : new()
    {
        return ParseDocument<T>(fileStream, worksheetNumber, false);
    }

    /// <inheritdoc />
    public List<T> ParseDocument<T>(Stream fileStream, int worksheetNumber, bool skipHeaderRow) where T : new()
    {
        //Validate object is properly created
        var importColumnDefinitions = typeof(T)
            .GetProperties()
            .Where(x => x.CustomAttributes.Any(c => c.AttributeType == typeof(SpreadsheetImportColumnAttribute)))
            .Select(p => new
            {
                Property = p,
                Column = p.GetCustomAttributes<SpreadsheetImportColumnAttribute>().First()
                    .ColumnIndex //safe because if where above
            }).ToList();

        if (importColumnDefinitions.Count == 0)
            throw new ArgumentException("No columns identified as SpreadsheetImportColumns, unable to process", "T");

        //Import
        var excelDoc = SpreadsheetDocument.Open(fileStream, false);
        var workbookPart = excelDoc.WorkbookPart;
        if (workbookPart == null) throw new SpreadsheetParserException("Spreadsheet has no WorkbookPart");

        var sheet = workbookPart.Workbook.Descendants<Sheet>().Skip(worksheetNumber - 1).FirstOrDefault();
        if (sheet == null) throw new SpreadsheetParserException($"Workbook does not have {worksheetNumber} sheets");
        if (sheet.Id == null || !sheet.Id.HasValue || sheet.Id.Value == null) throw new SpreadsheetParserException($"Sheet {worksheetNumber} has a null Id");

        if (workbookPart.GetPartById(sheet.Id.Value) is not WorksheetPart wsPart) 
            throw new SpreadsheetParserException($"Sheet {worksheetNumber} with Id {sheet.Id.Value} is not in the workbook");
       
        var collection = new Collection<T>();
        var skipRows = skipHeaderRow ? 1 : 0;
        var expectedColumns = importColumnDefinitions.Max(c => c.Column) - 1;

        foreach (Row row in wsPart.Worksheet.Descendants<Row>().Skip(skipRows))
        {
            var tnew = new T();
            var cellCollection = row.Elements<Cell>().ToList();

            //Check to see if the row has at least the same number of cells as the import model expects.
            //If not, skip the row
            if (cellCollection.Count < expectedColumns)
                continue;

            foreach (var col in importColumnDefinitions)
            {
                if (cellCollection.ElementAtOrDefault(col.Column - 1) == null)
                {
                    continue;
                }

                var value = GetCellValue(cellCollection[col.Column - 1]);
                col.Property.SetValue(tnew, ValueFromCell(value, col.Property.PropertyType));
            }

            collection.Add(tnew);
        }


        return collection.ToList();

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

    private static DateTime? MangleDateTime(string value)
    {
        if (DateTime.TryParse(value, out var dt))
            return dt;
        if (double.TryParse(value, out var d))
            return DateTime.FromOADate(d);
        return null;
    }

    private static object? ValueFromCell(string? value, Type propertyType) => propertyType switch
    {
        _ when string.IsNullOrEmpty(value) => null,
        _ when IsOfType<int>(propertyType) => int.Parse(value),
        _ when IsOfType<decimal>(propertyType) => decimal.Parse(value),
        _ when IsOfType<double>(propertyType) => double.Parse(value),
        _ when IsOfType<long>(propertyType) => long.Parse(value),
        _ when IsOfType<float>(propertyType) => float.Parse(value),
        _ when IsOfType<DateTime>(propertyType) => MangleDateTime(value),
        _ when IsOfType<DateTimeOffset>(propertyType) => DateTimeOffset.Parse(value),
        _ => value
    };

    private static string? GetCellValue(Cell? cell)
    {
        if (cell == null)
            return null;
        if (cell.DataType == null)
            return cell.InnerText;

        string value = cell.InnerText;
        switch (cell.DataType.Value)
        {
            case CellValues.SharedString:
                // For shared strings, look up the value in the shared strings table.
                // Get worksheet from cell
                Debug.Assert(cell.Parent != null, "cell.Parent != null");
                OpenXmlElement parent = cell.Parent;
                while (parent.Parent != null && parent.Parent != parent
                                             && string.Compare(parent.LocalName, "worksheet", StringComparison.OrdinalIgnoreCase) != 0)
                {
                    parent = parent.Parent;
                }
                if (string.Compare(parent.LocalName, "worksheet", StringComparison.OrdinalIgnoreCase) != 0)
                {
                    throw new SpreadsheetParserException($"Unable to find parent worksheet of cell {cell}");
                }

                var ws = parent as Worksheet;
                var ssDoc = ws?.WorksheetPart?.OpenXmlPackage as SpreadsheetDocument;
                var sstPart = ssDoc?.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                return sstPart == null ? value : sstPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
            //this case within a case is copied from msdn. 
            case CellValues.Boolean:
                return value switch
                {
                    "0" => "FALSE",
                    _ => "TRUE"
                };
            case CellValues.Number:
            case CellValues.Error:
            case CellValues.String:
            case CellValues.InlineString:
            case CellValues.Date:
            default:
                return value;
        }
    }
}