# NetCore.Utilities.Spreadsheet ![](https://img.shields.io/github/license/iowacomputergurus/netcore.utilities.spreadsheet.svg)

![Build Status](https://github.com/IowaComputerGurus/netcore.utilities.spreadsheet/actions/workflows/ci-build.yml/badge.svg)

A utility to assist in creating Excel spreadsheets in .NET Core and ASP.NET Core applications using the OpenXML library.  This utility allows you to export collections of .NET Objects to Excel by simply adding metadata information regarding the desired column formats, title etc.  Allowing quick & consistent excel exports, without the hassle of trying to understand the OpenXML format

## NuGet Package Information
ICG.NetCore.Utilities.Spreadsheet ![](https://img.shields.io/nuget/v/icg.netcore.utilities.spreadsheet.svg) ![](https://img.shields.io/nuget/dt/icg.netcore.utilities.spreadsheet.svg)

## SonarCloud Analysis

[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=IowaComputerGurus_netcore.utilities.spreadsheet&metric=alert_status)](https://sonarcloud.io/dashboard?id=IowaComputerGurus_netcore.utilities.spreadsheet)
[![Coverage](https://sonarcloud.io/api/project_badges/measure?project=IowaComputerGurus_netcore.utilities.spreadsheet&metric=coverage)](https://sonarcloud.io/dashboard?id=IowaComputerGurus_netcore.utilities.spreadsheet)
[![Security Rating](https://sonarcloud.io/api/project_badges/measure?project=IowaComputerGurus_netcore.utilities.spreadsheet&metric=security_rating)](https://sonarcloud.io/dashboard?id=IowaComputerGurus_netcore.utilities.spreadsheet)
[![Technical Debt](https://sonarcloud.io/api/project_badges/measure?project=IowaComputerGurus_netcore.utilities.spreadsheet&metric=sqale_index)](https://sonarcloud.io/dashboard?id=IowaComputerGurus_netcore.utilities.spreadsheet)

## Dependencies
This project depends on the DocumentFormat.OpenXml NuGet package provided by the Microsoft team. It is a MIT licensed library.

## Usage

## Installation
Standard installation via NuGet Package Manager
``` powershell
Install-Package ICG.NetCore.Utilities.Spreadsheet
```

## Setup
To setup the needed dependency injection items for this library, add the following line in your DI setup.
``` csharp
services.UseIcgNetCoreUtilitiesSpreadsheet();
```

## Sample Single Document Export

Exporting a single collection to a single excel file can be done very simply. 

```csharp
var exportGenerator = provider.GetService<ISpreadsheetGenerator>();
var exportDefinition = new SpreadsheetConfiguration<SimpleExportData>
{
    RenderTitle = true,
    DocumentTitle = "Sample Export of 100 Records",
    RenderSubTitle = true,
    DocumentSubTitle = "Showing the full options",
    ExportData = GetSampleExportData(100),
    WorksheetName = "Sample",
    FreezePanes = true,
    AutoFilterDataRows = true
};
var fileContent = exportGenerator.CreateSingleSheetSpreadsheet(exportDefinition);
System.IO.File.WriteAllBytes("Sample.xlsx", fileContent);
```

## Sample Multi-Sheet Document Export

A streamlined fluent syntax is available to export multiple sheets of content.

```csharp
var multiSheetDefinition = new MultisheetConfiguration()
    .WithSheet("Sheet 1", GetSampleExportData(100))
    .WithSheet("Additional Sheet", GetSampleExportData(500), config =>
    {
        config.DocumentTitle = "Lots of data";
        config.RenderTitle = true;
    });

var multiFileContent = exportGenerator.CreateMultiSheetSpreadsheet(multiSheetDefinition);
System.IO.File.WriteAllBytes("Sample-Multi.xlsx", multiFileContent);
```

## Key Features
This package is primarily geared towards the exporting of lists of objects into excel sheets.  The following key features are supported.

* The ability to have one, or more, sheets of data exported
* The ability to have a heading and subheading if desired
* Data type formatting for Date & Currency fields
* Auto-fit of all columns for display
* The ability to freeze the header columns into a freeze pane for single sheet, or multi-sheet exports
* The ability to add "Auto Filter" behavior to the data table portion of a sheet, while still supporting all other items
* The ability to automatically add "simple formula" totals to columsn. (SUM, AVG, etc)
* Support for Curreny, Date, F0, F1, F2, and F3 fixed data formats