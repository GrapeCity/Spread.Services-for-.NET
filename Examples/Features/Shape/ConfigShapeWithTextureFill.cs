using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Shape
{
    public class ConfigShapeWithTextureFill : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddShape(GrapeCity.Documents.Spread.Drawing.AutoShapeType.Heart, 1, 1, 100, 100);
            shape.Fill.PresetTextured(GrapeCity.Documents.Spread.Drawing.PresetTexture.Canvas);
            shape.Fill.TextureAlignment = GrapeCity.Documents.Spread.Drawing.TextureAlignment.Center;
            shape.Fill.TextureOffsetX = 2.5;
            shape.Fill.TextureOffsetY = 3.2;
            shape.Fill.TextureHorizontalScale = 0.9;
            shape.Fill.TextureVerticalScale = 0.2;
            shape.Fill.Transparency = 0.5;
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
