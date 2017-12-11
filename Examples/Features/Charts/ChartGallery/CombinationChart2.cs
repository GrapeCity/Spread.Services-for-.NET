using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.ChartGallery
{
    public class CombinationChart2 : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            worksheet.Range["A1:C17"].Value = new object[,] {
                { "Area 1", "Column 1", "Column 2" },
                { 1350, 120, 75 },
                { 1500, 90, 35 },
                { 1200, 80, 50 },
                { 1300, 80, 80 },
                { 1750, 90, 100 },
                { 1640, 120, 130 },
                { 1700, 120, 95 },
                { 1100, 90, 80 },
                { 1350, 120, 75 },
                { 1500, 90, 35 },
                { 1200, 80, 50 },
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:C17"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Combination Chart";
            GrapeCity.Documents.Spread.Drawing.ISeries series1 = shape.Chart.SeriesCollection[0];
            GrapeCity.Documents.Spread.Drawing.ISeries series2 = shape.Chart.SeriesCollection[1];
            GrapeCity.Documents.Spread.Drawing.ISeries series3 = shape.Chart.SeriesCollection[2];
            //change series type
            series1.ChartType = GrapeCity.Documents.Spread.Drawing.ChartType.Area;
            series2.ChartType = GrapeCity.Documents.Spread.Drawing.ChartType.ColumnStacked;
            series3.ChartType = GrapeCity.Documents.Spread.Drawing.ChartType.ColumnStacked;
            //set axis group
            series2.AxisGroup = GrapeCity.Documents.Spread.Drawing.AxisGroup.Secondary;
            series3.AxisGroup = GrapeCity.Documents.Spread.Drawing.AxisGroup.Secondary;
            //config axis sacle and unit
            GrapeCity.Documents.Spread.Drawing.IAxis value_axis = shape.Chart.Axes.Item(GrapeCity.Documents.Spread.Drawing.AxisType.Value);
            GrapeCity.Documents.Spread.Drawing.IAxis value_second_axis = shape.Chart.Axes.Item(GrapeCity.Documents.Spread.Drawing.AxisType.Value, GrapeCity.Documents.Spread.Drawing.AxisGroup.Secondary);
            value_axis.MaximumScale = 1800;
            value_axis.MajorUnit = 450;
            value_second_axis.MaximumScale = 300;
            value_second_axis.MajorUnit = 75;
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
