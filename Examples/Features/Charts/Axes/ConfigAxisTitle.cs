﻿using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Charts.Axes
{
    public class ConfigAxisTitle : ExampleBase
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

            category_axis.HasTitle = true;
            category_axis.AxisTitle.Format.Fill.Color.RGB = Color.Pink;
            category_axis.AxisTitle.Text = "category axis";
            category_axis.AxisTitle.Font.Size = 10;
            category_axis.AxisTitle.Font.Color.RGB = Color.Green;
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
