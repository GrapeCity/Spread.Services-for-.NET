using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Shape
{
    public class ConfigShapeWithGradientFill : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddShape(GrapeCity.Documents.Spread.Drawing.AutoShapeType.Heart, 1, 1, 100, 100);
            shape.Fill.PresetGradient(GrapeCity.Documents.Spread.Drawing.GradientStyle.Vertical, 3, GrapeCity.Documents.Spread.Drawing.PresetGradientType.Silver);
            shape.Fill.RotateWithObject = false;
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
