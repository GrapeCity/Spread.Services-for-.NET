using GrapeCity.Documents.Spread.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.PageSetup
{
    public class ConfigFirstPageHeaderFooter : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream("PageSetup Demo.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set first page headerfooter
            worksheet.PageSetup.DifferentFirstPageHeaderFooter = true;

            worksheet.PageSetup.FirstPage.CenterHeader.Text = "&T";
            worksheet.PageSetup.FirstPage.RightFooter.Text = "&D";

            //Set first page headerfooter's graphic
            worksheet.PageSetup.FirstPage.LeftFooter.Text = "&G";
            var assembly = this.GetType().GetTypeInfo().Assembly;
            System.IO.Stream stream = assembly.GetManifestResourceStream("GrapeCity.Documents.Spread.Examples.Resource.logo.png");
            worksheet.PageSetup.FirstPage.LeftFooter.Picture.SetGraphicStream(stream, ImageType.PNG);

        }

        public override bool IsNew
        {
            get
            {
                return true;
            }
        }
    }
}
