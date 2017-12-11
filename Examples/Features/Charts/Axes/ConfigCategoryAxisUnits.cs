using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.Axes
{
    public class ConfigCategoryAxisUnits : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.Range["A2:A6"].NumberFormat = "m/d/yyyy";
            worksheet.Range["A1:D6"].Value = new object[,]
            {
                {null, "S1", "S2", "S3"},
                {new DateTime(2015, 10, 7), 10, 25, 25},
                {new DateTime(2015, 10, 24), 51, 36, 27},
                {new DateTime(2015, 11, 8), 52, 85, 30},
                {new DateTime(2015, 11, 25), 22, 65, 65},
                {new DateTime(2015, 12, 10), 23, 69, 69}
            };

            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns, true, true);

            GrapeCity.Documents.Spread.Drawing.IAxis category_axis = shape.Chart.Axes.Item(GrapeCity.Documents.Spread.Drawing.AxisType.Category);
            category_axis.MaximumScale = new DateTime(2015, 12, 20).ToOADate();
            category_axis.MinimumScale = new DateTime(2015, 10, 1).ToOADate();
            category_axis.BaseUnit = GrapeCity.Documents.Spread.Drawing.TimeUnit.Months;
            category_axis.MajorUnitScale = GrapeCity.Documents.Spread.Drawing.TimeUnit.Months;
            category_axis.MajorUnit = 1;
            category_axis.MinorUnitScale = GrapeCity.Documents.Spread.Drawing.TimeUnit.Days;
            category_axis.MinorUnit = 15;
        }


    }
}
