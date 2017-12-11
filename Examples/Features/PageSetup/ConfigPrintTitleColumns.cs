using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.PageSetup
{
    public class ConfigPrintTitleColumns : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream("PageSetup Print Title Columns.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set columns to repeat at left
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
