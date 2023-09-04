using System;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            Sections();
        }

		/// <summary>
        /// Creates a document with different sections.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/sections-and-page-layout.php
        /// </remarks>
        static void Sections()
        {
            string documentPath = @"Sections.docx";

            // Let's create a simple document with two Sections.
            DocumentCore dc = new DocumentCore();

            // First Section, A4 - Portrait.
            Section s1 = new Section(dc);
            s1.PageSetup.PaperType = PaperType.A4;
            s1.PageSetup.Orientation = Orientation.Portrait;
            dc.Sections.Add(s1);

            // Add some text into section 1.
            s1.Content.Start.Insert("Text in section 1", new CharacterFormat() { FontName = "Times New Roman", Size = 60.0 });

            // Second Section, Letter - Landscape.
            Section s2 = new Section(dc);
            s2.PageSetup.PaperType = PaperType.Letter;
            s2.PageSetup.Orientation = Orientation.Landscape;
            dc.Sections.Add(s2);

            // Add some text into section 2.
            s2.Content.Start.Insert("Text in section 2", new CharacterFormat() { FontName = "Arial", Size = 72.0, FontColor = new Color("DD55AA") });

            // Save our document into DOCX format.
            dc.Save(documentPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}