using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            StyleInheritance();
        }
        /// <summary>
        /// How the Styles Inheritance does work.
        /// </summary>
        /// <remarks>
        /// https://sautinsoft.com/products/document/help/net/developer-guide/styles-inheritance.php
        /// </remarks>
        static void StyleInheritance()
        {
            string docxPath = @"StylesInheritance.docx";

            // Let's create document.
            DocumentCore dc = new DocumentCore();
            dc.DefaultCharacterFormat.FontColor = Color.Blue; 
            Section section = new Section(dc);
            section.Blocks.Add(new Paragraph(dc, new Run(dc, "The document has Default Character Format with 'Blue' color.", new CharacterFormat() { Size = 18 })));
            dc.Sections.Add(section);
            
            // Create a new Paragraph and Style with 'Yellow' background.
            Paragraph par = new Paragraph(dc);
            ParagraphStyle styleYellowBg = new ParagraphStyle("YellowBackground");
            styleYellowBg.CharacterFormat.BackgroundColor = Color.Yellow;
            dc.Styles.Add(styleYellowBg);
            par.ParagraphFormat.Style = styleYellowBg;

            par.Inlines.Add(new Run(dc, "This paragraph has Style 'Yellow Background' and it inherits 'Blue Color' from the document's DefaultCharacterFormat."));
            par.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.LineBreak));
            Run run1 =  new Run(dc, "This Run doesn't have a style, but it inherits 'Yellow Background' from the paragraph style and 'Blue Color' from the document's DefaultCharacterFormat.");
            run1.CharacterFormat.Italic = true;
            par.Inlines.Add(run1);
            par.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.LineBreak));

            Run run2 = new Run(dc, " This run has own Style with 'Green Color'.");
            CharacterStyle styleGreenText = new CharacterStyle("GreenText");
            styleGreenText.CharacterFormat.FontColor = Color.Green;
            dc.Styles.Add(styleGreenText);
            run2.CharacterFormat.Style = styleGreenText;            
            par.Inlines.Add(run2);

            Paragraph par2 = new Paragraph(dc);
            Run run3 = new Run(dc, "This is a new paragraph without a style. This is a Run also without style. " +
                "But they both inherit 'Blue Color' from their parent - the document.");
            par2.Inlines.Add(run3);
            section.Blocks.Add(par);
            section.Blocks.Add(par2);

            // Save our document into DOCX format.
            dc.Save(docxPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docxPath) { UseShellExecute = true });
        }
    }
}