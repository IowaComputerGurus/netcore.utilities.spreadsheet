using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ICG.NetCore.Utilities.Spreadsheet
{
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
                    Column = p.GetCustomAttributes<SpreadsheetImportColumnAttribute>().First().ColumnIndex //safe because if where above
                }).ToList();

            if (importColumnDefinitions.Count == 0)
                throw new ArgumentException("No columns identified as SpreadsheetImportColumns, unable to process", "T");

            //Import
            using (fileStream)
            {
                var excelDoc = SpreadsheetDocument.Open(fileStream, false);
                var workbookPart = excelDoc.WorkbookPart;
                var sheet = excelDoc.WorkbookPart.Workbook.Descendants<Sheet>().ToList()[worksheetNumber - 1]; //Offset due to 1 based values
                var wsPart = workbookPart.GetPartById(sheet.Id) as WorksheetPart;
                var collection = new Collection<T>();
                var skipRows = skipHeaderRow ? 1 : 0;
                foreach (Row row in wsPart.Worksheet.Descendants<Row>().Skip(skipRows))
                {
                    var tnew = new T();
                    var cellCollection = row.Elements<Cell>().ToList();

                    foreach (var col in importColumnDefinitions)
                    {
                        var value = GetCellValue(cellCollection[col.Column - 1]);
                        if (string.IsNullOrEmpty(value))
                        {
                            col.Property.SetValue(tnew, null);
                        }

                        else if (col.Property.PropertyType == typeof(int))
                        {
                            col.Property.SetValue(tnew, int.Parse(value));
                        }

                        else if (col.Property.PropertyType == typeof(double))
                        {
                            col.Property.SetValue(tnew, double.Parse(value));
                        }

                        else if (col.Property.PropertyType == typeof(DateTime))
                        {
                            col.Property.SetValue(tnew, DateTime.Parse(value));
                        }
                        else
                        {
                            //Its a string
                            col.Property.SetValue(tnew, value);
                        }
                    }

                    collection.Add(tnew);
                }

                return collection.ToList();
            }
        }

        private static string GetCellValue(Cell cell)
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
                    OpenXmlElement parent = cell.Parent;
                    while (parent.Parent != null && parent.Parent != parent
                                                 && string.Compare(parent.LocalName, "worksheet", true) != 0)
                    {
                        parent = parent.Parent;
                    }
                    if (string.Compare(parent.LocalName, "worksheet", true) != 0)
                    {
                        throw new Exception("Unable to find parent worksheet.");
                    }

                    Worksheet ws = parent as Worksheet;
                    SpreadsheetDocument ssDoc = ws.WorksheetPart.OpenXmlPackage as SpreadsheetDocument;
                    SharedStringTablePart sstPart = ssDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    // lookup value in shared string table
                    if (sstPart != null && sstPart.SharedStringTable != null)
                    {
                        value = sstPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                    }
                    break;

                //this case within a case is copied from msdn. 
                case CellValues.Boolean:
                    switch (value)
                    {
                        case "0":
                            value = "FALSE";
                            break;
                        default:
                            value = "TRUE";
                            break;
                    }
                    break;
            }
            return value;
        }
    }
}
