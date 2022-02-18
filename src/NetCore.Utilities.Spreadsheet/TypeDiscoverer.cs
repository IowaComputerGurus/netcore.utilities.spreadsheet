using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ICG.NetCore.Utilities.Spreadsheet
{
    internal record PropDetail(
        int Order,
        PropertyDescriptor Descriptor,
        string DisplayName,
        string Format,
        float Width
    );

    /// <summary>
    /// Probes a type to get information that is used to structure the spreadsheet
    /// </summary>
    internal static class TypeDiscoverer
    {
        private static readonly Regex TypeNameRegex = new(@"(?<=[A-Z])(?=[A-Z][a-z])|(?<=[^A-Z])(?=[A-Z])|(?<=[A-Za-z])(?=[^A-Za-z])");
        public static IList<PropDetail> GetProps(Type t)
        {
            var properties = TypeDescriptor.GetProperties(t);
            var details = new List<PropDetail>();
            var columnOrder = 1;
            var width = 0f;
            var format = "";
            foreach (PropertyDescriptor p in properties)
            {
                var propName = p.DisplayName;
                if (p.DisplayName == p.Name)
                {
                    propName = TypeNameRegex.Replace(p.Name, " ");
                }

                var ignored = false;
                foreach (var attr in p.Attributes)
                {
                    switch (attr)
                    {
                        case SpreadsheetIgnoreAttribute _:
                            // Since we're ignoring this column, don't need to do anything else
                            ignored = true;
                            continue;
                            break;
                        case SpreadsheetColumnFormatAttribute cfa:
                            format = cfa.Format;
                            break;
                        case SpreadsheetColumnAttribute sca:
                            if (sca.Ignore)
                            {
                                ignored = true;
                                continue;
                            }
                            format = sca.Format ?? format;
                            propName = sca.DisplayName ?? propName;
                            width = sca.Width;
                            break;
                    }
                }

                if (ignored) continue;

                details.Add(new PropDetail(columnOrder, p, propName, format, width));
                columnOrder++;
            }

            return details;
        }
    }
}
