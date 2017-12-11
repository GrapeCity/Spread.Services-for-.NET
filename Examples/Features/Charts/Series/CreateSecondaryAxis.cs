using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.Series
{
    public class CreateSecondaryAxis:ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            worksheet.Range["A1:C6"].Value = new object[,]
            {
                { null, "S1", "S2"},
                { "Item1", 10, 25},
                { "Item2", -51, -36},
                { "Item3", 32, 64},
                { "Item4", 44, 80},
                { "Item5", 60,100}
            };

            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:D6"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns, true, true);
            GrapeCity.Documents.Spread.Drawing.ISeries series2 = shape.Chart.SeriesCollection[1];
            //add a secondary axis
            series2.AxisGroup = GrapeCity.Documents.Spread.Drawing.AxisGroup.Secondary;
            series2.ChartType = GrapeCity.Documents.Spread.Drawing.ChartType.Line;
          
        }
    }
}
