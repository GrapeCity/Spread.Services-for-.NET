using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Workbook
{
    public class CreateNewWorkbook : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            //Create empty workbook, contains one worksheet default.
            GrapeCity.Documents.Spread.Workbook workbookNew = new GrapeCity.Documents.Spread.Workbook();
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
