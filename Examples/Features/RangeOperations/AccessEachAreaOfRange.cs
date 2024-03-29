﻿using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.RangeOperations
{
    public class AccessEachAreaOfRange : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            var range = worksheet.Range["A5:B7, C3, H5:N6"];

            //set interior color for area1 A5:B7.
            var area1 = worksheet.Range["A5:B7, C3, H5:N6"].Areas[0];
            area1.Interior.Color = Color.Pink;

            //set interior color for area2 C3.
            var area2 = worksheet.Range["A5:B7, C3, H5:N6"].Areas[1];
            area2.Interior.Color = Color.LightGreen;

            //set interior color for area3 H5:N6.
            var area3 = worksheet.Range["A5:B7, C3, H5:N6"].Areas[2];
            area3.Interior.Color = Color.LightBlue;
        }
    }
}
