using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.ChartGallery
{
    public class XYScatterLinesChart:ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.XYScatterLines, 250, 20, 360, 230);
            worksheet.Range["A1:B8"].Value = new object[,] {
                { 75, 250 },
                { 50, 125 },
                { 25, 375 },
                { 75, 250 },
                { 50, 875 },
                { 25, 625 },
                { 75, 750 },
                { 125, 500 },
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B8"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Scatter with Straight Lines and Markers Chart";
        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
