﻿using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.PageSetup
{
    public class ConfigHeaderFooter : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            var fileStream = this.GetTemplateStream("PageSetup Demo.xlsx");
            workbook.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set page headerfooter
            worksheet.PageSetup.LeftHeader = "&\"Arial,Italic\"LeftHeader";
            worksheet.PageSetup.CenterHeader = "&P";

            //Set page headerfooter's graphic
            worksheet.PageSetup.CenterFooter = "&G";
            var assembly = this.GetType().GetTypeInfo().Assembly;
            System.IO.Stream stream = assembly.GetManifestResourceStream("GrapeCity.Documents.Spread.Examples.Resource.logo.png");
            worksheet.PageSetup.CenterFooterPicture.SetGraphicStream(stream, Drawing.ImageType.PNG);

            //If you have picture resources locally, you can also set graphic in this way.
            //worksheet.PageSetup.CenterFooter = "&G";
            //worksheet.PageSetup.CenterFooterPicture.Filename = @"C:\picture.png";

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
