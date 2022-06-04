using SoftUniTableOfContentOOP.Interfaces;
using Syncfusion.DocIO.DLS;
using System;
using System.Drawing;
using BuiltinStyle = Syncfusion.DocIO.DLS.BuiltinStyle;
using HorizontalAlignment = Syncfusion.DocIO.DLS.HorizontalAlignment;

namespace SoftUniTableOfContentOOP
{
    public class MainClass : IAdd
    {
        public WParagraph AddParagraph(WSection section, BuiltinStyle builtinStyle, string headingText, string paragraghText)
        {
            WParagraph newPara = (WParagraph)section.AddParagraph();

            newPara.AppendText(headingText);

            if (headingText.Contains("Section"))
            {
                newPara.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;
            }

            newPara.ApplyStyle(builtinStyle);
            newPara = (WParagraph)section.AddParagraph();
            newPara.AppendText(paragraghText);
            section.AddParagraph();

            return newPara;
        }

        public void AddTitle(WordDocument document, string text)
        {
            WParagraph titlePara = document.LastParagraph;
            titlePara.AppendText(text);

            //Setting the font of the title...
            WParagraphStyle style = document.Styles.FindByName("Normal") as WParagraphStyle;
            style.CharacterFormat.FontName = "Arial";
            style.CharacterFormat.FontSize = 20;
            style.Name = "MyStyle";
            style.CharacterFormat.Bold = true;
            titlePara.ApplyStyle("MyStyle");
            titlePara.ParagraphFormat.HorizontalAlignment = HorizontalAlignment.Center;

            titlePara = (WParagraph)document.LastSection.AddParagraph();
            TableOfContent toc = titlePara.AppendTOC(1, 4);
        }

        public BuiltinStyle AddHeading(int num)
        {
            if (num == 1) { return BuiltinStyle.Heading1; }
            else if (num == 2) { return BuiltinStyle.Heading2; }
            else if (num == 3) { return BuiltinStyle.Heading3; }
            else if (num == 4) { return BuiltinStyle.Heading4; }
            else if (num == 5) { return BuiltinStyle.Heading5; }
            else { throw new Exception("Invalid heading!"); }
        }



    }
}
