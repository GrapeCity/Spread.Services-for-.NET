using System;
using System.Collections.Generic;
using System.Text;

namespace GrapeCity.Documents.Spread.Examples.Features.Theme
{
    public class ChangeWorkbookTheme : ExampleBase
    {
        public override void Execute(GrapeCity.Documents.Spread.Workbook workbook)
        {
            //Change workbook's theme to built-in theme.
            workbook.Theme = Themes.Berlin;

            Themes themes = new Themes();
            //base theme is office theme when not give parameter.
            ITheme theme = themes.Add("testTheme");

            //Change workbook's theme to custom theme.
            workbook.Theme = theme;

            //Customize theme's color.
            theme.ThemeColorScheme[ThemeColor.Light1].RGB = Color.AntiqueWhite;
            theme.ThemeColorScheme[ThemeColor.Accent1].RGB = Color.AliceBlue;
            //Customize theme's font.
            theme.ThemeFontScheme.Major[FontLanguageIndex.Latin].Name = "Buxton Sketch";
            theme.ThemeFontScheme.Minor[FontLanguageIndex.Latin].Name = "Segoe UI";
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
