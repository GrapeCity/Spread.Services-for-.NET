﻿using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Filtering.FilteringByIcon
{
    public class FilterByNoIcon : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            IWorksheet worksheet = workbook.Worksheets[0];

            object[,] data = new object[,]{
                {"Name", "City", "Birthday", "Sex", "Weight", "Height"},
                {"Bob", "NewYork", new DateTime(1968, 6, 8), "male", 80, 180},
                {"Betty", "NewYork", new DateTime(1972, 7, 3), "female", 72, 168},
                {"Gary", "NewYork", new DateTime(1964, 3, 2), "male", 71, 179},
                {"Hunk", "Washington", new DateTime(1972, 8, 8), "male", 80, 171},
                {"Cherry", "Washington", new DateTime(1986, 2, 2), "female", 58, 161},
                { "Eva", "Washington", new DateTime(1993, 2, 15), "female", 71, 180}
            };

            worksheet.Range["A1:F7"].Value = data;
            worksheet.Range["A:F"].ColumnWidth = 15;

            IIconSetCondition iconset = worksheet.Range["E2:E7"].FormatConditions.AddIconSetCondition();
            iconset.IconSet = workbook.IconSets[IconSetType.Icon3TrafficLights1];

            IFormatCondition condition = worksheet.Range["E2:E7"].FormatConditions.Add(FormatConditionType.CellValue, FormatConditionOperator.Between, "71", "79") as IFormatCondition;
            condition.StopIfTrue = true;

            //filter in the rows which has no icon.
            worksheet.Range["A1:F7"].AutoFilter(4, null, AutoFilterOperator.NoIcon);
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
