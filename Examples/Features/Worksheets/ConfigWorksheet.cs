using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Worksheets
{
    public class ConfigWorksheet : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //worksheet sheet view settings.
            worksheet.SheetView.DisplayFormulas = false;
            worksheet.SheetView.DisplayRightToLeft = true;
            
            //Set worksheet tab color.
            worksheet.TabColor = Color.Green;

            //Set worksheet default row height.
            worksheet.StandardHeight = 20;
            //Set worksheet default column width.
            worksheet.StandardWidth = 50;

            IWorksheet worksheet1 = workbook.Worksheets.Add();
            //Hide worksheet.
            worksheet1.Visible = Visibility.Hidden;
        }

        public override bool ShowViewer
        {
            get
            {
                return false;
            }
        }

        public override bool ShowScreenshot
        {
            get
            {
                return true;
            }
        }

    }
}
