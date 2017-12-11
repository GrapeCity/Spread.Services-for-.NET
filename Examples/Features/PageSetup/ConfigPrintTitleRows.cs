using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.PageSetup
{
    public class ConfigPrintTitleRows : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream("PageSetup Print Title Rows.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set rows to repeat at top
            worksheet.PageSetup.PrintTitleRows = "$5:$10";

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
