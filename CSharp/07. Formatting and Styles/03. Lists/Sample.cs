using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
       
        static void Main(string[] args)
        {
            SimpleLists();
        }

		/// <summary>
        /// How to create a simple document with ordered and unordered lists. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/lists.php
        /// </remarks>		
        public static void SimpleLists()
        {
            string documentPath = @"Lists.pdf";

            // Let's create a simple document.
            DocumentCore dc = new DocumentCore();

            // Add a new section.
            Section s = new Section(dc);
            dc.Sections.Add(s);

            string[] myCollection = new string[] { "One", "Two", "Three", "Four", "Five" };

            // Create new ordered list style.
            ListStyle ordList = new ListStyle("Simple Numbers", ListTemplateType.NumberWithDot);
            foreach (ListLevelFormat f in ordList.ListLevelFormats)
            {
                f.CharacterFormat.Size = 14;
            }
            dc.Styles.Add(ordList);

            // Add caption
            dc.Content.Start.Insert("This is a simple ordered list:", new CharacterFormat() { Size = 14.0, FontColor = Color.Black });

            // Add the collection of paragraphs marked as ordered list.
            int level = 0;
            foreach (string listText in myCollection)
            {
                Paragraph p = new Paragraph(dc);
                dc.Sections[0].Blocks.Add(p);

                p.Content.End.Insert(listText, new CharacterFormat() { Size = 14.0, FontColor = Color.Black });
                p.ListFormat.Style = ordList;
                p.ListFormat.ListLevelNumber = level;
                p.ParagraphFormat.SpaceAfter = 0;
            }

            // Add the collection of paragraphs marked as unordered list (bullets).

            // Add caption
            dc.Content.End.Insert("\n\nThis is a simple bulleted list:", new CharacterFormat() { Size = 14.0, FontColor = Color.Black });

            // Create list style.
            ListStyle bullList = new ListStyle("Bullets", ListTemplateType.Bullet);
            dc.Styles.Add(bullList);

            level = 0;
            foreach (string listText in myCollection)
            {
                Paragraph p = new Paragraph(dc);
                dc.Sections[0].Blocks.Add(p);

                p.Content.End.Insert(listText, new CharacterFormat() { Size = 14.0, FontColor = Color.Black });
                p.ListFormat.Style = bullList;
                p.ListFormat.ListLevelNumber = level;
                p.ParagraphFormat.SpaceAfter = 0;
            }

            // Save our document into PDF format.
            dc.Save(documentPath, new PdfSaveOptions() { Compliance = PdfCompliance.PDF_A1a });

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}