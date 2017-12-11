using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.PageSetup
{
    public class ConfigFirstPageNumber : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream("PageSetup Demo.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.PageSetup.CenterFooter = "&P";

            //Set first page number, default is p1.
            worksheet.PageSetup.FirstPageNumber = 3;

        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
