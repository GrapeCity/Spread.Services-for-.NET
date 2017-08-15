# Spread.Services-for-.Net
GrapeCity **Spread.Services for .NET** is a new small-footprint high-performance spreadsheet component. It gives developers a comprehensive API to quickly create, manipulate, convert, and share Microsoft Excel-compatible spreadsheets. Plus, you can call it from nearly any application and platform.
 
We model our interface-based API on Excel's document object model. This means that you can import, calculate, query, generate, and export any spreadsheet scenario.
 
Imported or generated spreadsheets can contain references to one another. Some of the many things you can reference include: full reports, sorted and filtered tables, sorted and filtered pivot tables, charts, sparklines, conditional formats, and dashboard reports.
 
**Spread.Services** targets .NET Standard 1.4 for multi-platform support including the .NET Framework, .NET Core, Mono and more. **Spread.Services** is the perfect solution for all of your spreadsheet challenges.

This repository contains source project of Examples and Showcases of **Spread.Services** to help you learn and write your own applications. Note that you need to install [.Net Core SDK](https://www.microsoft.com/net/core) to run these examples.

| Directory    | Description    |
| ------------- |-------------|
| Examples     | A collection of .NET examples that help you learn and explore the API features |
| AspNetCoreDemo/AspNetCore+React     | A source project that demonstrates how to use Spread.Services with Asp .Net Core + React + Spread.Sheets, to run this project, besides [.Net Core SDK](https://www.microsoft.com/net/core), you also need to install [NodeJS](https://nodejs.org/en/) |
| AspNetCoreDemo/AspNetCore+Angular2     | A source project that demonstrates how to use Spread.Services with Asp .Net Core + Angular2 + Spread.Sheets,to run this project, besides [.Net Core SDK](https://www.microsoft.com/net/core), you also need to install [NodeJS](https://nodejs.org/en/)|
| Benchmark | Contains source projects to help users run performance for spread.services, put the excel files into the Files\Input folder, then run the project you will get the performance data.|

# Limitations of non-licensed package
The Spread.Services package we use in all of these projects are not licensed, there will be some limiations for the non-licensed package:
* You can just open or save 100 Excel files
* You can not create new instance of **Workbook** after your programme runs after 10 hours
* Exceptions will be thrown when you call some of Spread.Services api after your programme runs after 10 hours
* A new worksheet with watermark will be added and activated everytime when you save to an excel file

If you want to use Spread.Services without any limiations, you can visit http://spread.grapecity.com or email spread.sales@grapecity.com to get an free evaluation license, or purchase a production license for distribution.

# Resource
* Online Demo: [http://spread.cloudapp.net/spreadservices](http://spread.cloudapp.net/spreadservices)
* Product Home Site: [http://spread.grapecity.com](http://spread.grapecity.com)
* Nuget Package: [https://www.nuget.org/packages/Spread.Services/](https://www.nuget.org/packages/Spread.Services/)
* .Net Core SDK: [https://www.microsoft.com/net/core](https://www.microsoft.com/net/core)
