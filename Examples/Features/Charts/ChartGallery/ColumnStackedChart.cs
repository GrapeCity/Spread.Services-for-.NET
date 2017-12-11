using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.ChartGallery
{
    public class ColumnStackedChart : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.ColumnStacked, 250, 20, 360, 230);
            worksheet.Range["A1:C6"].Value = new object[,] {
                { 103, 121, 109 },
                { 56, 94, 115 },
                { 116, 89, 99 },
                { 55, 93, 70 },
                { 114, 114, 83 },
                { 125, 138, 136 },
            };

            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:C6"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns);
            shape.Chart.ChartTitle.Text = "Column Stacked Chart";
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
