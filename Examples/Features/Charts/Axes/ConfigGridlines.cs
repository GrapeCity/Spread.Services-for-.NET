using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.Axes
{
    public class ConfigGridlines : ExampleBase
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
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns, true, true);

            GrapeCity.Documents.Spread.Drawing.IAxis category_axis = shape.Chart.Axes.Item(GrapeCity.Documents.Spread.Drawing.AxisType.Category);
            category_axis.HasMajorGridlines = true;
            category_axis.HasMinorGridlines = true;
            category_axis.MajorGridlines.Format.Line.Color.RGB = Color.LightGreen;
            category_axis.MajorGridlines.Format.Line.Weight = 2;
            category_axis.MinorGridlines.Format.Line.Color.RGB = Color.Pink;
            category_axis.MinorGridlines.Format.Line.Weight = 1;
            category_axis.MinorGridlines.Format.Line.Style = GrapeCity.Documents.Spread.Drawing.LineStyle.ThickThin;
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
