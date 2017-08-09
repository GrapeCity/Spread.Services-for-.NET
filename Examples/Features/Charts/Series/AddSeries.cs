using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.Series
{
    public class AddSeries : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.ColumnClustered, 300, 10, 300, 300);
            worksheet.Range["A1:D6"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            };

            //Add series
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns, true, true);
            
            GrapeCity.Documents.Spread.Drawing.ISeries series1 = shape.Chart.SeriesCollection[0];
            GrapeCity.Documents.Spread.Drawing.ISeries series2 = shape.Chart.SeriesCollection[1];

            //change series1's axis group, plot on secondary axis.
            series1.AxisGroup = GrapeCity.Documents.Spread.Drawing.AxisGroup.Secondary;

            //change series2's chart type.
            series2.ChartType = GrapeCity.Documents.Spread.Drawing.ChartType.LineMarkers;

            //if it is a negative point, its color is red.
            series1.InvertIfNegative = true;
            series1.InvertColor.ObjectThemeColor = ThemeColor.Accent2;

            //config series2's marker style.
            series2.MarkerStyle = GrapeCity.Documents.Spread.Drawing.MarkerStyle.Diamond;
            series2.MarkerSize = 5;
            series2.MarkerFormat.Fill.Color.RGB = Color.Orange;
            series2.MarkerFormat.Line.Style = GrapeCity.Documents.Spread.Drawing.LineStyle.ThickThin;
            series2.MarkerFormat.Line.Color.RGB = Color.LightYellow;
            series2.MarkerFormat.Line.Weight = 3;

            //change series2's line color.
            series2.Format.Line.Color.RGB = Color.LightGreen;
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
