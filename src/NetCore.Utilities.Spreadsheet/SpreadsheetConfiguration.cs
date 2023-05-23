using System;
using System.Collections;
using System.Collections.Generic;

namespace ICG.NetCore.Utilities.Spreadsheet;

/// <summary>
///     Definition of a spreadsheet configuration for exporting of data
/// </summary>
public interface ISpreadsheetConfiguration
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
    IEnumerable ExportData { get; }

    /// <summary>
    /// The type of data represented in this export configuration
    /// </summary>
    Type DataType { get; }

    /// <summary>
    ///     The desired worksheet name for the export
    /// </summary>
    /// <value>The name of the worksheet.</value>
    string WorksheetName { get; set; }

    /// <summary>
    /// Controls if the columns are auto-sized
    /// </summary>
    public bool AutoSizeColumns { get; set; }

    /// <summary>
    /// if set to true the header(s) will be frozen
    /// </summary>
    public bool FreezeHeaders { get; set; }
}
/// <inheritdoc />
/// <typeparam name="TRecord">The type to be exported</typeparam>
public interface ISpreadsheetConfiguration<out TRecord> : ISpreadsheetConfiguration where TRecord : class
{
    /// <inheritdoc cref="ExportData"/>
    new IEnumerable<TRecord> ExportData { get; }

    IEnumerable ISpreadsheetConfiguration.ExportData => ExportData;
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

    /// <inheritdoc cref="ISpreadsheetConfiguration{T}"/>
    public Type DataType => typeof(T);

    /// <summary>
    ///     The desired worksheet name for the export
    /// </summary>
    /// <value>The name of the worksheet.</value>
    public string WorksheetName { get; set; }


    /// <inheritdoc />
    public bool AutoSizeColumns { get; set; } = true;

    /// <summary>
    /// if set to true the header(s) will be frozen
    /// </summary>
    public bool FreezeHeaders { get; set; }
}

/// <summary>
///     Describes configuration for a multi-sheet export
/// </summary>
public class MultisheetConfiguration : IEnumerable<ISpreadsheetConfiguration>
{
    private readonly List<ISpreadsheetConfiguration> _sheets = new();

    /// <summary>
    ///     Adds a sheet with data to the export
    /// </summary>
    /// <typeparam name="T">The type to be exported</typeparam>
    /// <param name="worksheetName">The name of the worksheet.</param>
    /// <param name="data">The data for the worksheet</param>
    /// <returns>
    ///     The same instance of <see cref="MultisheetConfiguration"/> to allow for
    ///     for fluent configuration
    /// </returns>
    public MultisheetConfiguration WithSheet<T>(string worksheetName, IEnumerable<T> data) where T : class
    {
        _sheets.Add(new SpreadsheetConfiguration<T>{ WorksheetName = worksheetName, ExportData = data});
        return this;
    }

    /// <inheritdoc cref="WithSheet{T}(string,IEnumerable{T})"/>
    /// <param name="worksheetName">The name of the worksheet.</param>
    /// <param name="data">The data for the worksheet</param>
    /// <param name="config">A callback allowing for additional configuration of the sheet</param>
    public MultisheetConfiguration WithSheet<T>(string worksheetName, IEnumerable<T> data, Action<SpreadsheetConfiguration<T>> config) where T : class
    {
        var sheet = new SpreadsheetConfiguration<T> { WorksheetName = worksheetName, ExportData = data };
        config(sheet);
        _sheets.Add(sheet);
        return this;
    }

    /// <inheritdoc />
    public IEnumerator<ISpreadsheetConfiguration> GetEnumerator() => _sheets.GetEnumerator();
    IEnumerator IEnumerable.GetEnumerator() => ((IEnumerable)_sheets).GetEnumerator();
}