# NetCore.Utilities.Spreadsheet ![](https://img.shields.io/github/license/iowacomputergurus/netcore.utilities.spreadsheet.svg)
A utility to assist in creating Excel spreadsheets in .NET Core and ASP.NET Core applications using the OpenXML library.  This utility allows you to export collections of
.NET Objects to Excel by simply adding metadata information regarding the desired column formats, title etc.  Allowing quick & consistent excel exports, without the hassle of trying
to understand the OpenXML format

## NuGet Package Information
ICG.NetCore.Utilities.Spreadsheet ![](https://img.shields.io/nuget/v/icg.netcore.utilities.spreadsheet.svg) ![](https://img.shields.io/nuget/dt/icg.netcore.utilities.spreadsheet.svg)

## Dependencies
This project depends on the DocumentFormat.OpenXml NuGet package provided by the Microsoft team. It is a MIT licensed library.

## Usage

## Installation
Standard installation via NuGet Package Manager
```
Install-Package ICG.NetCore.Utilities.Spreadsheet
```

## Setup
To setup the needed dependency injection items for this library, add the following line in your DI setup.
```
services.UseIcgNetCoreUtilitiesSpreadsheet();
```

## Sample Single Document Export

Exporting a single collection to a single excel file can be done very simply. 

```
var exportGenerator = provider.GetService<ISpreadsheetGenerator>();
var exportDefinition = new SpreadsheetConfiguration<SimpleExportData>
{
    RenderTitle = true,
    DocumentTitle = "Sample Export of 100 Records",
    RenderSubTitle = true,
    DocumentSubTitle = "Showing the full options",
    ExportData = GetSampleExportData(100),
    WorksheetName = "Sample"
};
var fileContent = exportGenerator.CreateSingleSheetSpreadsheet(exportDefinition);
System.IO.File.WriteAllBytes("Sample.xlsx", fileContent);
```

## Key Features
This package is primarily geared towards the exporting of lists of objects into excel sheets.  The following key features are supported.

* The ability to have one, or more, sheets of data exported
* The ability to have a heading and subheading if desired
* Data type formatting for Date & Currency fields
* Auto-fit of all columns for display