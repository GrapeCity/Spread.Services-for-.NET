using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.RangeOperations
{
    public class SetRowHeightColumnWidth : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            //set row height for row 1:2.
            worksheet.Range["1:2"].RowHeight = 50;

            //set column width for column C:D.
            worksheet.Range["C:D"].ColumnWidth = 20;
        }
    }
}
