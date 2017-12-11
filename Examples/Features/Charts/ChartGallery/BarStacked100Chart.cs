using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.ChartGallery
{
    public class BarStacked100Chart : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.BarStacked100, 250, 20, 360, 230);
            worksheet.Range["A1:B5"].Value = new object[,] {
                { 1, 5 },
                { 2, 4 },
                { 3, 3 },
                { 4, 2 },
                { 4, 1 }
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B5"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Bar Stacked 100 Chart";
            shape.Chart.Legend.Position = GrapeCity.Documents.Spread.Drawing.LegendPosition.Left;
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
