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
                {new DateTime(2015, 10, 21), 10, 25, 25},
                {new DateTime(2016, 10, 25), -51, -36, 27},
                {new DateTime(2017, 12, 20), 52, -85, -30},
                {new DateTime(2018, 5, 5), 22, 65, 65},
                {new DateTime(2019, 10, 12), 23, 69, 69}
            };

            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.ColumnClustered, 300, 10, 300, 300);
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns, true, true);

            GrapeCity.Documents.Spread.Drawing.IAxis category_axis = shape.Chart.Axes.Item(GrapeCity.Documents.Spread.Drawing.AxisType.Category);

            category_axis.MaximumScale = new DateTime(2019, 10, 1).ToOADate();
            category_axis.MinimumScale = new DateTime(2015, 10, 1).ToOADate();
            category_axis.BaseUnit = GrapeCity.Documents.Spread.Drawing.TimeUnit.Years;
            category_axis.MajorUnitScale = GrapeCity.Documents.Spread.Drawing.TimeUnit.Months;
            category_axis.MajorUnit = 4;
            category_axis.MinorUnitScale = GrapeCity.Documents.Spread.Drawing.TimeUnit.Days;
            category_axis.MinorUnit = 60;
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
