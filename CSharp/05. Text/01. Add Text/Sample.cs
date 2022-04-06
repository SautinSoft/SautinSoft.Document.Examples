using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            AddText();
        }

        /// <summary>
        /// How to create a simple document with text. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/text.php
        /// </remarks>
        public static void AddText()
        {
            string documentPath = @"Text.docx";

            // Let's create a new document.
            DocumentCore dc = new DocumentCore();

            // Create a new section and add into the document.
            Section section = new Section(dc);
            dc.Sections.Add(section);

            // Create a new paragraph and add into the section.
            Paragraph par = new Paragraph(dc);
            section.Blocks.Add(par);

            // Create Inline-derived objects with text.
            Run run1 = new Run(dc, "This is a rich");
            run1.CharacterFormat = new CharacterFormat() { FontName = "Times New Roman", Size = 18.0, FontColor = new Color(112, 173, 71) };

            Run run2 = new Run(dc, " formatted text");
            run2.CharacterFormat = new CharacterFormat() { FontName = "Arial", Size = 10.0, FontColor = new Color("#0070C0") };

            SpecialCharacter spch3 = new SpecialCharacter(dc, SpecialCharacterType.LineBreak);

            Run run4 = new Run(dc, "with a line break.");
            run4.CharacterFormat = new CharacterFormat() { FontName = "Times New Roman", Size = 10.0, FontColor = Color.Black };

            // Add our inlines into the paragraph.
            par.Inlines.Add(run1);
            par.Inlines.Add(run2);
            par.Inlines.Add(spch3);
            par.Inlines.Add(run4);

            // Save our document into the DOCX format.
            dc.Save(documentPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}