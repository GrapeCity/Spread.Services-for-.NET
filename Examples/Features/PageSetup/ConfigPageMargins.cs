using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.PageSetup
{
    public class ConfigPageMargins : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream("PageSetup Demo.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set margins, in points.
            worksheet.PageSetup.TopMargin = 36;
            worksheet.PageSetup.BottomMargin = 36;
            worksheet.PageSetup.RightMargin = 72;

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
