﻿using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.DataValidation
{
    public class CreateTextLength : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["C2:E4"].Value = new object[,]
            {
                {"aa", "bb", "aa1"},
                {"ccc", "dddd", "dddd1"},
                {"ff", "ffff", "ffff1"}
            };

            //create text length validation, text length between 2 and 3.
            worksheet.Range["C2:E4"].Validation.Add(ValidationType.TextLength, ValidationAlertStyle.Stop, ValidationOperator.Between, 2, 3);

            //judge if Range["C2:E4"] has validation.
            for (int i = 1; i <= 3; i++)
            {
                for (int j = 2; j <= 4; j++)
                {
                    if (worksheet.Range[i, j].HasValidation)
                    {
                        //set the range[i, j]'s interior color.
                        worksheet.Range[i, j].Interior.Color = Color.LightBlue;
                    }
                }
            }
        }
    }
}
