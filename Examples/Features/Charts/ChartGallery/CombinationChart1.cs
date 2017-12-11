using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.ChartGallery
{
    public class CombinationChart1 :ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            worksheet.Range["A1:B13"].Value = new object[,] {
                { "Blue Column", "Red Line" },
                { 75, 20 },
                { 149, 50 },
                { 105, 30 },
                { 55, 80 },
                { 121, 40 },
                { 76, 110 },
                { 128, 50 },
                { 114, 140 },
                { 75, 60 },
                { 105, 170 },
                { 145, 70 },
                { 110, 100 },
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B13"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Combination Chart";
            //change series type
            GrapeCity.Documents.Spread.Drawing.ISeries series2 = shape.Chart.SeriesCollection[1];
            series2.ChartType = GrapeCity.Documents.Spread.Drawing.ChartType.LineMarkers;
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
