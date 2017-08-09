using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Worksheets
{
    public class ActivateWorksheet : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets.Add();
            //Activate new created worksheet.
            worksheet.Activate();
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }
    }
}
