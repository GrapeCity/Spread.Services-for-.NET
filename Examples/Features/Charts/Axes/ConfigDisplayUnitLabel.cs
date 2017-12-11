using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.Axes
{
    public class ConfigDisplayUnitLabel : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            worksheet.Range["A1:D6"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {"Item1", 10, 25, 25},
                {"Item2", -51, -36, 27},
                {"Item3", 52, -85, -30},
                {"Item4", 22, 65, 65},
                {"Item5", 23, 69, 69}
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns, true, true);

            GrapeCity.Documents.Spread.Drawing.IAxis value_axis = shape.Chart.Axes.Item(GrapeCity.Documents.Spread.Drawing.AxisType.Value);
            value_axis.DisplayUnit = GrapeCity.Documents.Spread.Drawing.DisplayUnit.Custom;
            value_axis.DisplayUnitCustom = 100;
            value_axis.HasDisplayUnitLabel = true;
            value_axis.DisplayUnitLabel.Font.Color.RGB = Color.CornflowerBlue;
            value_axis.DisplayUnitLabel.Format.Fill.Color.RGB = Color.Orange;
            value_axis.DisplayUnitLabel.Format.Line.Color.RGB = Color.CornflowerBlue;
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
