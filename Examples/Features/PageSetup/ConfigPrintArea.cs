using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.PageSetup
{
    public class ConfigPrintArea : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream("PageSetup Print Area.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set print area & print titles
            worksheet.PageSetup.PrintArea = "$D$5:$G$10";
            worksheet.PageSetup.PrintTitleRows = "$5:$10";
            worksheet.PageSetup.PrintTitleColumns = "$D:$G";

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
