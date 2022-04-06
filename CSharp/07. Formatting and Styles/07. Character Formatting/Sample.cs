using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
       
        static void Main(string[] args)
        {
            CharacterFormatting();
        }

		/// <summary>
        /// This sample shows how to set character format. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/character-format.php
        /// </remarks>		
        public static void CharacterFormatting()
        {
            string documentPath = @"CharacterFormat.pdf";

            // Let's create a simple document.
            DocumentCore dc = new DocumentCore();

            // Add a new section.
            dc.Sections.Add(new Section(dc));

            // Add a paragraph.
            Paragraph p = new Paragraph(dc);
            p.ParagraphFormat.Alignment = HorizontalAlignment.Left;
            dc.Sections[0].Blocks.Add(p);

            // Create a formatted text (Run element) and add it into paragraph.
            Run run1 = new Run(dc, "It\'s wide formatted text.");
            run1.CharacterFormat.AllCaps = true;
            run1.CharacterFormat.BackgroundColor = Color.Pink;
            run1.CharacterFormat.FontName = "Verdana";
            run1.CharacterFormat.Size = 26f;
            run1.CharacterFormat.FontColor = new Color("#FFFFFF");

            p.Inlines.Add(run1);

            // Create another Run element (container for characters).
            Run run2 = new Run(dc, "Hi from SautinSoft!");
            run2.CharacterFormat.FontColor = Color.DarkGreen;
            run2.CharacterFormat.UnderlineStyle = UnderlineType.Dashed;
            run2.CharacterFormat.UnderlineColor = Color.Gray;

            // Add another formatted text into the paragraph.
            p.Inlines.Add(run2);

            // Add new paragraph with formatted text.
            // We are using ContentRange to insert text.
            dc.Content.Start.Insert("This is the first paragraph.\n", new CharacterFormat() { FontName = "Calibri", Size = 16.0, FontColor = Color.Orange, Bold = true });
            (dc.Sections[0].Blocks[0] as Paragraph).ParagraphFormat.Alignment = HorizontalAlignment.Center;

            dc.Content.End.Insert("Bold", new CharacterFormat() { Bold = true, FontName = "Times New Roman", Size = 11.0 });
            dc.Content.End.Insert(" Italic ", new CharacterFormat() { Italic = true, FontName = "Calibri", Size = 11.0 });
            dc.Content.End.Insert("Underline", new CharacterFormat() { UnderlineStyle = UnderlineType.Single, FontName = "Calibri", Size = 11.0 });
            dc.Content.End.Insert(" ", new CharacterFormat() { Bold = true, FontName = "Segoe UI", Size = 11.0 });
            dc.Content.End.Insert("Strikethrough", new CharacterFormat() { Strikethrough = true, FontName = "Calibri", Size = 11.0 });

            // Save our document into PDF format.
            dc.Save(documentPath, new PdfSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}