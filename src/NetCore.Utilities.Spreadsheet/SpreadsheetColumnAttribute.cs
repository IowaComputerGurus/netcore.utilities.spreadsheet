using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ICG.NetCore.Utilities.Spreadsheet
{
    /// <summary>
    /// Controls how a property is mapped to a spreadsheet column
    /// </summary>
    public class SpreadsheetColumnAttribute : Attribute
    {
        public string DisplayName { get; }
        public float Width { get; }
        public bool Ignore { get; }
        public string Format { get; }

        /// <summary>
        /// Initializes a new 
        /// </summary>
        /// <param name="displayName">Sets the display name of the column. If not provided, will fall back on the DisplayName attribute.</param>
        /// <param name="width">Sets the width of the column</param>
        /// <param name="ignore">If true, the column will be excluded from the spreadsheet</param>
        /// <param name="format">Sets the format of the column data</param>
        public SpreadsheetColumnAttribute(string displayName = null, float width = 0, bool ignore = false, string format = null)
        {
            DisplayName = displayName;
            Width = width;
            Ignore = ignore;
            Format = format;

        }
    }
}
