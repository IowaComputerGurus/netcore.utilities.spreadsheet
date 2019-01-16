using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;

namespace ICG.NetCore.Utilities.Spreadsheet
{
    /// <inheritdoc />
    public class EPPlusSpreadsheetParser : ISpreadsheetParser
    {
        /// <inheritdoc />
        public List<T> ParseDocument<T>(Stream fileStream) where T : new()
        {
            return ParseDocument<T>(fileStream, 1);
        }

        /// <inheritdoc />
        public List<T> ParseDocument<T>(Stream fileStream, int worksheetNumber) where T : new()
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

            if(importColumnDefinitions.Count == 0) 
                throw new ArgumentException("No columns identified as SpreadsheetImportColumns, unable to process", "T");

            //Import
            using (fileStream)
            {
                var excel = new ExcelPackage(fileStream);
                using (var worksheet = excel.Workbook.Worksheets[worksheetNumber])
                {
                    //Get the count of rows
                    var rows = worksheet.Cells
                        .Select(cell => cell.Start.Row)
                        .OrderBy(x => x).ToList();

                    var collection = new Collection<T>();

                    for (var i = 1; i < rows.Count; i++)
                    {
                        var row = rows[i];
                        var tnew = new T();
                        foreach (var col in importColumnDefinitions)
                        {
                            //This is the real wrinkle to using reflection - Excel stores all numbers as double including int
                            var val = worksheet.Cells[row, col.Column];
                            //If it is numeric it is a double since that is how excel stores all numbers
                            if (val.Value == null)
                            {
                                col.Property.SetValue(tnew, null);
                            }

                            else if (col.Property.PropertyType == typeof(int))
                            {
                                col.Property.SetValue(tnew, val.GetValue<int>());
                            }

                            else if (col.Property.PropertyType == typeof(double))
                            {
                                col.Property.SetValue(tnew, val.GetValue<double>());
                            }

                            else if (col.Property.PropertyType == typeof(DateTime))
                            {
                                col.Property.SetValue(tnew, val.GetValue<DateTime>());
                            }
                            else
                            {
                                //Its a string
                                col.Property.SetValue(tnew, val.GetValue<string>());
                            }
                        }
                        collection.Add(tnew);
                    }

                    return collection.ToList();
                }
            }
        }
    }
}
