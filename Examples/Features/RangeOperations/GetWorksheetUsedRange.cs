﻿using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.RangeOperations
{
    public class GetWorksheetUsedRange : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["H6:M7"].Value = 1;
            worksheet.Range["J9:J10"].Merge();

            //set interior color for worksheet usedRange "H6:M10".
            var usedrange = worksheet.UsedRange;
            usedrange.Interior.Color = Color.LightBlue;
        }
    }
}
