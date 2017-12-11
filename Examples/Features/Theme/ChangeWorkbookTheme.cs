using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Theme
{
    public class ChangeWorkbookTheme : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            //Change workbook's theme to built-in theme.
            workbook.Theme = Themes.Berlin;

            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["B2"].Value = "Text";

        }
    }
}
