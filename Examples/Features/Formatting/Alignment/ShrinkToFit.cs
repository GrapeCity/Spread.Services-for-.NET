﻿using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Formatting.Alignment
{
    public class ShrinkToFit : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];
            IRange rangeB1 = worksheet.Range["B1"];
            rangeB1.Value = "The ShrinkToFit property is applied";
            rangeB1.ShrinkToFit = true;
        }
    }
}
