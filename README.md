# Spread.Services-for-.NET
GrapeCity **Spread.Services for .NET** is a new small-footprint high-performance spreadsheet component that can be used in your server or desktop applications. It gives developers a comprehensive API to quickly create, manipulate, convert, and share Microsoft Excel-compatible spreadsheets. Plus, you can call it from nearly any application and platform. 

We model our interface-based API on Excel's document object model. This means that you can import, calculate, query, generate, and export any spreadsheet scenario. 

Imported or generated spreadsheets can contain references to one another. Some of the many things you can reference include: full reports, sorted and filtered tables, sorted and filtered pivot tables, charts, sparklines, conditional formats, and dashboard reports. 

**Spread.Services** targets multiple platforms including the .NET Framework, .NET Core and Mono. Spread.Services is the perfect solution for all of your spreadsheet challenges.

This repository contains source project of Examples and Showcases of **Spread.Services** to help you learn and write your own applications. Note that you need to install [.NET Core SDK](https://www.microsoft.com/net/core) to run these examples.

| Directory    | Description    |
| ------------- |-------------|
| Examples     | A collection of .NET examples that help you learn and explore the API features |
| AspNetCoreDemo/AspNetCore+React     | A source project that demonstrates how to use Spread.Services with Asp.NET Core + React + Spread.Sheets (to run this project, install [.NET Core SDK](https://www.microsoft.com/net/core) and [NodeJS](https://nodejs.org/en/)) |
| AspNetCoreDemo/AspNetCore+Angular2     | A source project that demonstrates how to use Spread.Services with Asp .Net Core + Angular2 + Spread.Sheets(to run this project, install [.NET Core SDK](https://www.microsoft.com/net/core) and [NodeJS](https://nodejs.org/en/))|
| Benchmark | Source projects to help users run performance tests on Spread.Services (Put Excel files into the Files\Input folder and run the project to get performance data)|

# Limitations of non-licensed package
These projects use the non-licensed version of Spread.Services. The non-licensed version has the following limitations:
* You can only open or save 100 Excel files.
* You can only run an application for up to 10 hours
* When you save a file, a new worksheet with watermark will be added stating that this was generated using a non-licensed evaluation version.

# Other Resources
* Online Demo: [http://spread.cloudapp.net/spreadservices](http://spread.cloudapp.net/spreadservices)
* Product Home Site: [http://spread.grapecity.com](http://spread.grapecity.com)
* Nuget Package: [https://www.nuget.org/packages/Spread.Services/](https://www.nuget.org/packages/Spread.Services/)
* .NET Core SDK: [https://www.microsoft.com/net/core](https://www.microsoft.com/net/core)
