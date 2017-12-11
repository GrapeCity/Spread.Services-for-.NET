using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;

namespace GrapeCity.Documents.Spread.Examples.Features.Shape
{
    public class ConfigShapeWithPictureFill : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddShape(GrapeCity.Documents.Spread.Drawing.AutoShapeType.Rectangle, 20, 20, 100, 100);

            var assembly = this.GetType().GetTypeInfo().Assembly;
            System.IO.Stream stream = assembly.GetManifestResourceStream("GrapeCity.Documents.Spread.Examples.Resource.logo.png");

            shape.Fill.UserPicture(stream, GrapeCity.Documents.Spread.Drawing.ImageType.JPG);
            shape.PictureFormat.ColorType = GrapeCity.Documents.Spread.Drawing.PictureColorType.Grayscale;
            shape.PictureFormat.Brightness = 0.6;
            shape.PictureFormat.Contrast = 0.3;
            shape.PictureFormat.Crop.PictureOffsetX = 10;
            shape.PictureFormat.Crop.PictureOffsetY = -5;
            shape.PictureFormat.Crop.PictureWidth = 120;
            shape.PictureFormat.Crop.PictureHeight = 80;
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
