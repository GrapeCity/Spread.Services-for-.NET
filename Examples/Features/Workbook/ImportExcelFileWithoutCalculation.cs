using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Workbook
{
    public class ImportExcelFileWithoutCalculation : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            // When OpenOptions.DoNotRecalculateAfterOpened means Spread.Services will just read all the cached values without calculating again after
            // opened an Excel file
            //change the path to the real file path when open.
            workbook.Open(this.CurrentDirectory + "source.xlsx", null, new OpenOptions() { DoNotRecalculateAfterOpened = true });
        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }

        public override bool CanDownload
        {
            get
            {
                return false;
            }
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
