using SautinSoft.Document;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            Styles();
        }

        /// <summary>
        /// This sample shows how to work with styles. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/styles.php
        /// </remarks>
        public static void Styles()
        {
            string docxPath = @"Styles.docx";

            // Let's create document.
            DocumentCore dc = new DocumentCore();

            // Create custom styles.
            CharacterStyle characterStyle = new CharacterStyle("CharacterStyle1");
            characterStyle.CharacterFormat.FontName = "Arial";
            characterStyle.CharacterFormat.UnderlineStyle = UnderlineType.Wave;
            characterStyle.CharacterFormat.Size = 18;

            ParagraphStyle paragraphStyle = new ParagraphStyle("ParagraphStyle1");
            paragraphStyle.CharacterFormat.FontName = "Times New Roman";
            paragraphStyle.CharacterFormat.Size = 14;
            paragraphStyle.ParagraphFormat.Alignment = HorizontalAlignment.Center;

            // Add styles to the document, then use it.
            dc.Styles.Add(characterStyle);
            dc.Styles.Add(paragraphStyle);

            // Add text content.
            dc.Sections.Add(
            new Section(dc,
                new Paragraph(dc,
                new Run(dc, "Once upon a time, in a far away swamp, there lived an ogre named "),
                new Run(dc, "Shrek") { CharacterFormat = { Style = characterStyle } },
                new Run(dc, " whose precious solitude is suddenly shattered by an invasion of annoying fairy tale characters..."))
                { ParagraphFormat = { Style = paragraphStyle } }));

            // Save our document into DOCX format.
            dc.Save(docxPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docxPath) { UseShellExecute = true });
        }
    }
}