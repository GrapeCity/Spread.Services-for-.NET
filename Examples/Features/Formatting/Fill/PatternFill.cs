using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Formatting.Fill
{
    public class PatternFill : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A1"].Interior.Pattern = GrapeCity.Documents.Spread.Pattern.LightDown;
            worksheet.Range["A1"].Interior.Color = Color.Pink;
            worksheet.Range["A1"].Interior.PatternColorIndex = 5;
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
