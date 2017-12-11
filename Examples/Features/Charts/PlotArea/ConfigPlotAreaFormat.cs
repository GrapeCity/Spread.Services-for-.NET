using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.PlotArea
{
    public class ConfigPlotAreaFormat : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            worksheet.Range["A1:D6"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, 36, 27},
                {"Item3", 52, 50, -30},
                {"Item4", 22, 65, 30},
                {"Item5", 23, 40, 69}
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns, true, true);

            GrapeCity.Documents.Spread.Drawing.IPlotArea plotarea = shape.Chart.PlotArea;
            plotarea.Format.Fill.Color.RGB = Color.LightGray;
            plotarea.Format.Line.Color.RGB = Color.Gray;
            plotarea.Format.Line.Weight = 1;
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
