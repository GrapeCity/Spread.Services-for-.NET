using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.Axes
{
    public class ConfigTickMarkLabel : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            worksheet.Range["A1:D6"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {1, -25, 25, 25},
                {2, 51, 36, 27},
                {3, 52, 80, 30},
                {4, 22, -20, 65},
                {5, 23, 69, 69}
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns, true, true);
            GrapeCity.Documents.Spread.Drawing.IAxis category_axis = shape.Chart.Axes.Item(GrapeCity.Documents.Spread.Drawing.AxisType.Category);
            GrapeCity.Documents.Spread.Drawing.IAxis value_axis = shape.Chart.Axes.Item(GrapeCity.Documents.Spread.Drawing.AxisType.Value);

            //config tick label's format
            category_axis.TickLabelPosition = GrapeCity.Documents.Spread.Drawing.TickLabelPosition.NextToAxis;
            category_axis.TickLabelSpacing = 2;
            category_axis.TickLabels.Font.Color.RGB = Color.DarkOrange;           
            category_axis.TickLabels.Font.Size = 12;
            category_axis.TickLabels.NumberFormat = "#,##0.00";
            value_axis.TickLabels.NumberFormat = "#,##0;[Red]#,##0";
        }
    }
}
