using Syncfusion.DocIO.DLS;

namespace SoftUniTableOfContentOOP.Interfaces
{
    public interface IAdd
    {
        public WParagraph AddParagraph(WSection section, BuiltinStyle builtinStyle, string headingText, string paragraghText);
        public void AddTitle(WordDocument document, string text);
        public BuiltinStyle AddHeading(int num);
    }
}
