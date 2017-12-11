using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.ChartGallery
{
   public class LineChart :ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            GrapeCity.Documents.Spread.Drawing.IShape shape = worksheet.Shapes.AddChart(GrapeCity.Documents.Spread.Drawing.ChartType.Line, 250, 20, 360, 230);
            worksheet.Range["A1:C7"].Value = new object[,] {
                { 0, 59.18, 27.14 },
                { 44.64, 52.22, 25.08 },
                { 45.21, 49.80, 57.99 },
                { 24.32, 37.30, 42.73 },
                { 58.34, 34.43, 28.34 },
                { 31.89, 69.78, 46.88 },
                { 41.79, 63.94, 56.24 },   
            };           
            shape.Chart.SeriesCollection.Add(worksheet.Range["A1:C7"], GrapeCity.Documents.Spread.Drawing.RowCol.Columns);
            //set series lines style
            shape.Chart.SeriesCollection[0].Format.Line.Weight = 2.25;
            shape.Chart.SeriesCollection[1].Format.Line.Weight = 2.25;
            shape.Chart.SeriesCollection[2].Format.Line.Weight = 2.25;
            shape.Chart.ChartTitle.Text = "Line Chart";           
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
