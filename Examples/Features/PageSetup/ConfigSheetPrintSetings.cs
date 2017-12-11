using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.PageSetup
{
    public class ConfigSheetPrintSetings : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream("PageSetup Demo.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set sheet
            worksheet.PageSetup.PrintGridlines = true;
            worksheet.PageSetup.PrintHeadings = true;
            worksheet.PageSetup.BlackAndWhite = true;
            worksheet.PageSetup.PrintComments = PrintLocation.InPlace;
            worksheet.PageSetup.PrintErrors = PrintErrors.Dash;

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
