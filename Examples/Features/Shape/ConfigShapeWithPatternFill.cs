using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Shape
{
    public class ConfigShapeWithPatternFill : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddShape(GrapeCity.Documents.Spread.Drawing.AutoShapeType.Heart, 1, 1, 100, 100);
            shape.Fill.Patterned(GrapeCity.Documents.Spread.Drawing.PatternType.Percent10);
            shape.Fill.Color.ObjectThemeColor = ThemeColor.Accent2;
            shape.Fill.PatternColor.ObjectThemeColor = ThemeColor.Accent6;
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
