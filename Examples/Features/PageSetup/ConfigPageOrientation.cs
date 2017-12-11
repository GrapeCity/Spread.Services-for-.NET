using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.PageSetup
{
    public class ConfigPageOrientation : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream("PageSetup Demo.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set page orientation, default is portrait.
            worksheet.PageSetup.Orientation = PageOrientation.Landscape;

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
