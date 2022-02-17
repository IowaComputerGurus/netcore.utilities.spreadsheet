using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ICG.NetCore.Utilities.Spreadsheet
{
    
    /// <summary>
    /// Marks a property to be ignored when exporting to a spreadsheet
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class SpreadsheetIgnoreAttribute :Attribute
    {
    }
}
