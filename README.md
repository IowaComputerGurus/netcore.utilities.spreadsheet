# NetCore.Utilities.Spreadsheet ![](https://img.shields.io/github/license/iowacomputergurus/netcore.utilities.spreadsheet.svg)
A utility to assist in creating Excel spreadsheets in .NET Core and ASP.NET Core applications using the EPPlus.Core library.  This utility allows you to export .NET Object to Excel by simply adding metadata information regarding the desired column formats, title etc.  Allowing quick & consistent excel exports.

## NuGet Package Information
ICG.NetCore.Utilities.Spreadsheet ![](https://img.shields.io/nuget/v/icg.netcore.utilities.spreadsheet.svg) ![](https://img.shields.io/nuget/dt/icg.netcore.utilities.spreadsheet.svg)

## Dependencies
This project depends on the EPPlus.Core NuGet package.  No changes are made to the EPPlus.Core package, and its usage is goverened by its own license agreement.

## Usage

## Installation
Standard installation via HuGet Package Manager
```
Install-Package ICG.NetCore.Utilities.Spreadsheet
```

## Setup
To setup the needed dependency injection items for this library, add the following line in your DI setup.
```
services.UseIcgNetCoreUtilitiesSpreadsheet();
```

## Creating Documents

We are continuing to update this information.  For the quickest getting started guide please review the "samples" directory for samples.