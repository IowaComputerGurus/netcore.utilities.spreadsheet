using System;
using System.Collections.Generic;

namespace ICG.NetCore.Utilities.Spreadsheet;

/// <summary>
///     Definition of a spreadsheet configuration for exporting of data
/// </summary>
/// <typeparam name="T">The type to be exported</typeparam>
public interface ISpreadsheetConfiguration<out T> where T : class
{
    /// <summary>
    ///     Defines if the title will be rendered.  If [true] <see cref="DocumentTitle" /> is required.
    /// </summary>
    bool RenderTitle { get; set; }

    /// <summary>
    ///     Defines if the sub-title will be rendered.  If [true] <see cref="DocumentSubTitle" /> is required.
    /// </summary>
    bool RenderSubTitle { get; set; }

    /// <summary>
    ///     Top level heading to be supplied as part of the report
    /// </summary>
    /// <remarks>
    ///     Formatted as a 14pt bold header, if null, no header will render
    /// </remarks>
    /// <value>The document title.</value>
    string DocumentTitle { get; set; }

    /// <summary>
    ///     Secondary heading to be supplied as part of the report
    /// </summary>
    /// <remarks>
    ///     Placed in Row 3, 12pt bold header
    /// </remarks>
    /// <value>The document sub title.</value>
    string DocumentSubTitle { get; set; }

    /// <summary>
    ///     The data to export
    /// </summary>
    /// <value>The export data.</value>
    IEnumerable<T> ExportData { get; }

    /// <summary>
    ///     The type of data represented in this export configuration
    /// </summary>
    Type DataType { get; }

    /// <summary>
    ///     The desired worksheet name for the export
    /// </summary>
    /// <value>The name of the worksheet.</value>
    string WorksheetName { get; set; }

    /// <summary>
    ///     Controls if the columns are auto-sized
    /// </summary>
    public bool AutoSizeColumns { get; set; }
}

/// <summary>
///     Configuration class for passing export information to the spreadsheet operations
/// </summary>
public class SpreadsheetConfiguration<T> : ISpreadsheetConfiguration<T> where T : class
{
    /// <summary>
    ///     Defines if the title will be rendered.  If [true] <see cref="DocumentTitle" /> is required.
    /// </summary>
    public bool RenderTitle { get; set; }

    /// <summary>
    ///     Defines if the sub-title will be rendered.  If [true] <see cref="DocumentSubTitle" /> is required.
    /// </summary>
    public bool RenderSubTitle { get; set; }

    /// <summary>
    ///     Top level heading to be supplied as part of the report
    /// </summary>
    /// <remarks>
    ///     Formatted as a 14pt bold header, if null, no header will render
    /// </remarks>
    /// <value>The document title.</value>
    public string DocumentTitle { get; set; }

    /// <summary>
    ///     Secondary heading to be supplied as part of the report
    /// </summary>
    /// <remarks>
    ///     Placed in Row 3, 12pt bold header
    /// </remarks>
    /// <value>The document sub title.</value>
    public string DocumentSubTitle { get; set; }

    /// <summary>
    ///     The data to export
    /// </summary>
    /// <value>The export data.</value>
    public IEnumerable<T> ExportData { get; set; }

    /// <inheritdoc cref="ISpreadsheetConfiguration{T}" />
    public Type DataType => typeof(T);

    /// <summary>
    ///     The desired worksheet name for the export
    /// </summary>
    /// <value>The name of the worksheet.</value>
    public string WorksheetName { get; set; }


    /// <inheritdoc />
    public bool AutoSizeColumns { get; set; } = true;
}