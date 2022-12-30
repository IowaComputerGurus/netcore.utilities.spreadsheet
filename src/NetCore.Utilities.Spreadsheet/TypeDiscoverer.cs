using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace ICG.NetCore.Utilities.Spreadsheet;

internal record PropDetail(
    int Order,
    PropertyDescriptor Descriptor,
    string DisplayName,
    string Format,
    float Width
);

/// <summary>
///     Probes a type to get information that is used to structure the spreadsheet
/// </summary>
internal static class TypeDiscoverer
{
    private static readonly Regex TypeNameRegex = 
        new(@"(?<=[A-Z])(?=[A-Z][a-z])|(?<=[^A-Z])(?=[A-Z])|(?<=[A-Za-z])(?=[^A-Za-z])");

    public static IList<PropDetail> GetProps(Type t)
    {
        var properties = TypeDescriptor.GetProperties(t);
        var details = new List<PropDetail>();
        var columnOrder = 1;
        foreach (PropertyDescriptor p in properties)
        {
            var width = 0f;
            var format = "";
            var propName = p.DisplayName;
            if (p.DisplayName == p.Name) propName = TypeNameRegex.Replace(p.Name, " ");

            var ignored = false;
            foreach (var attr in p.Attributes)
            {
                if (attr is SpreadsheetColumnAttribute sca)
                {
                    if (sca.Ignore)
                    {
                        ignored = true;
                        continue;
                    }

                    format = (sca.Format ?? format).ToLowerInvariant();
                    propName = sca.DisplayName ?? propName;
                    width = sca.Width;
                }
                else if (attr is DisplayAttribute display)
                {
                    if (!string.IsNullOrEmpty(display.Name))
                        propName = display.Name;
                }
            }

            if (ignored) continue;

            details.Add(new PropDetail(columnOrder, p, propName, format, width));
            columnOrder++;
        }

        return details;
    }
}