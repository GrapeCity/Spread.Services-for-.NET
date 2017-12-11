using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.Series
{
    public class SetVaryColorForColumnChart : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.ColumnClustered, 250, 20, 360, 230);
            worksheet.Range["A1:B6"].Value = new object[,]
            {
                {null, "S1"},
                {"Item1", 10},
                {"Item2", -51},
                {"Item3", 52},
                {"Item4", 22},
                {"Item5", 23}
            };
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:B6"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns, true, true);
            //set vary colors for column chart which only has one series.
            shape.Chart.ColumnGroups[0].VaryByCategories = true;

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
