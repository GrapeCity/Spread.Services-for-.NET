using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;

namespace GrapeCity.Documents.Spread.Examples.Features.Picture
{
    public class DeletePicture : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            var assembly = this.GetType().GetTypeInfo().Assembly;
            System.IO.Stream stream = assembly.GetManifestResourceStream("GrapeCity.Documents.Spread.Examples.Resource.logo.png");
            GrapeCity.Documents.Spread.Drawing.IShape picture = worksheet.Shapes.AddPicture(stream, GrapeCity.Documents.Spread.Drawing.ImageType.PNG, 20, 20, 100, 100);
            //set picture size.
            picture.WidthInPixel = 700;
            picture.HeightInPixel = 120;
            //config picture layout.
            picture.Fill.Solid();
            picture.Fill.Color.ObjectThemeColor = ThemeColor.Accent1;

            //delete picture.
            picture.Delete();
        }
    }
}
