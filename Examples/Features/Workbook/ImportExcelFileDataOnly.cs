using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Workbook
{
    public class ImportExcelFileDataOnly : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            // Use OpenOptions.ImportFlags to control what you want to import from excel, ImportFlags.Data means only the data will be imported
            //change the path to the real file path when open.
            workbook.Open(this.CurrentDirectory + "source.xlsx", null, new OpenOptions() { ImportFlags = ImportFlags.Data });
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
