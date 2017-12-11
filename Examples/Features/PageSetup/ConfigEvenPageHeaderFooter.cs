using GrapeCity.Documents.Spread.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.PageSetup
{
    public class ConfigEvenPageHeaderFooter : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream("PageSetup Demo.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set even page headerfooter
            worksheet.PageSetup.OddAndEvenPagesHeaderFooter = true;

            worksheet.PageSetup.EvenPage.CenterHeader.Text = "&T";
            worksheet.PageSetup.EvenPage.RightFooter.Text = "&D";

            //Set even page headerfooter's graphic
            worksheet.PageSetup.EvenPage.LeftFooter.Text = "&G";
            var assembly = this.GetType().GetTypeInfo().Assembly;
            System.IO.Stream stream = assembly.GetManifestResourceStream("GrapeCity.Documents.Spread.Examples.Resource.logo.png");
            worksheet.PageSetup.EvenPage.LeftFooter.Picture.SetGraphicStream(stream, ImageType.PNG);

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
