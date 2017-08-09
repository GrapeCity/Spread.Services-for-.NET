using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Shape
{
    public class SetShapeText : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddShape(GrapeCity.Documents.Spread.Drawing.AutoShapeType.Heart, 1, 1, 100, 100);
            shape.Width = 600;
            shape.Height = 300;

            shape.TextFrame.TextRange.Font.Color.RGB = GrapeCity.Documents.Spread.Color.FromRGB(0, 255, 0);
            shape.TextFrame.TextRange.Font.Bold = true;
            shape.TextFrame.TextRange.Font.Italic = true;
            shape.TextFrame.TextRange.Font.Size = 20;
            shape.TextFrame.TextRange.Font.Strikethrough = true;

            shape.TextFrame.TextRange.Paragraphs.Add("This is a rectangle shape.");
            shape.TextFrame.TextRange.Paragraphs.Add("My name is xxx.");
            shape.TextFrame.TextRange.Paragraphs[1].Runs.Add("Hello World!");

            shape.TextFrame.TextRange.Paragraphs[1].Runs[0].Font.Strikethrough = false;
            shape.TextFrame.TextRange.Paragraphs[1].Runs[0].Font.Size = 35;
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
