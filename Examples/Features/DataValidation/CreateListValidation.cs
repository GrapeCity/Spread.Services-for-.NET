﻿using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.DataValidation
{
    public class CreateListValidation : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1"].Value = "aaa";
            worksheet.Range["A2"].Value = "bbb";
            worksheet.Range["A3"].Value = "ccc";

            worksheet.Range["C2:E4"].Value = new object[,]
            {
                {"aaa", "bbb", "ccc"},
                {"aaa1", "bbb1", "ccc1"},
                {"aaa2", "bbb2", "ccc2"}
            };

            //create list validation.
            worksheet.Range["C2:E4"].Validation.Add(ValidationType.List, ValidationAlertStyle.Stop, ValidationOperator.Between, "=$a$1:$a$3");
            IValidation validation = worksheet.Range["C2:E4"].Validation;
            validation.InCellDropdown = true;

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
