using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            AddHyperlink();
        }
        
		/// <summary>
        /// How to add a hyperlink into a document. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/hyperlinks.php
        /// </remarks>		
        public static void AddHyperlink()
        {
            string docxPath = @"Hyperlink.docx";

            // Let's create a simple DOCX document with a hyperlink.
            DocumentCore dc = new DocumentCore();

            Hyperlink hpl = new Hyperlink(dc, "http://www.zoo.org", "Welcome to Zoo!");
            (hpl.DisplayInlines[0] as Run).CharacterFormat = new CharacterFormat() { Size = 16, FontColor = new Color("#358CCB"), UnderlineStyle = UnderlineType.Single };
            hpl.ScreenTip = "Welcome to WoodLand Zoo!";

            Paragraph p = new Paragraph(dc);
            p.Inlines.Add(hpl);
            p.ParagraphFormat.Alignment = HorizontalAlignment.Center;

            dc.Content.Start.Insert(p.Content);

            // Save our document to DOCX format.
            dc.Save(docxPath);

           // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docxPath) { UseShellExecute = true });
        }
    }
}