using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Hyperlinks
{
    public class CreateHyperlinks : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A:A"].ColumnWidth = 20;
            //add a hyperlink link to web page.
            worksheet.Range["A1:B2"].Hyperlinks.Add(worksheet.Range["A1"],
                                                "http://www.baidu.com/",
                                                null,
                                                "open Baidu web site.",
                                                "Baidu");

            //add a hyperlink link to a range in this document.
            worksheet.Range["A3:B4"].Hyperlinks.Add(worksheet.Range["A3"],
                                            null,
                                            "Sheet1!$C$3:$E$4",
                                            "Go To sheet1 C3:E4");

            //add a hyperlink link to email address.
            worksheet.Range["A5:B6"].Hyperlinks.Add(worksheet.Range["A5"],
                                            "mailto:sophia.cao@grapecity.com",
                                            null,
                                            "Send an email to Sophia",
                                            "Send To Sophia");

            //add a hyperlink link to external file.
            //change the path to real picture file path.
            string path = this.CurrentDirectory + "external.xlsx";
            worksheet.Range["A7:B8"].Hyperlinks.Add(worksheet.Range["A7"],
                                                path,
                                                null,
                                                "link to external.xlsx file.",
                                                "external.xlsx");
        }
    }
}
