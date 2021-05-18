using ICG.NetCore.Utilities.Spreadsheet;

namespace Microsoft.Extensions.DependencyInjection
{
    /// <summary>
    /// Extension methods to make DI easier
    /// </summary>
    public static class StartupExtensions
    {
        /// <summary>
        ///     Registers the items included in the ICG AspNetCore Utilities project for Dependency Injection
        /// </summary>
        /// <param name="services">Your existing services collection</param>
        public static void UseIcgNetCoreUtilitiesSpreadsheet(this IServiceCollection services)
        {
            //Bind additional services
            services.AddTransient<ISpreadsheetGenerator, OpenXmlSpreadsheetGenerator>();
            services.AddTransient<ISpreadsheetParser, OpenXmlSpreadsheetParser>();
        }
    }
}