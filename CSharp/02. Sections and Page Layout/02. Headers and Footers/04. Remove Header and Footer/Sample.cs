using System.IO;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            ReplaceHeader();
        }

        /// <summary>
        /// Removes the old header/footer and inserts a new one into an existing PDF document.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/remove-header-and-footer-in-pdf-net-csharp-vb.php
        /// </remarks>
        static void ReplaceHeader()
        {
            string inpFile = @"..\..\..\somebody.pdf";
            string outFile = "With new Header.pdf";
            DocumentCore dc = DocumentCore.Load(inpFile);

            // Create new header with formatted text.
            HeaderFooter header = new HeaderFooter(dc, HeaderFooterType.HeaderDefault);
            header.Content.Start.Insert("Modified : 1 April 2020", new CharacterFormat() { Size = 14.0, FontColor = Color.DarkGreen });
            // Add 10 mm from Top before new header.
            (header.Blocks[0] as Paragraph).ParagraphFormat.SpaceBefore = LengthUnitConverter.Convert(10, LengthUnit.Millimeter, LengthUnit.Point);


            foreach (Section s in dc.Sections)
            {
                // Find the first paragraph (Let's assume that it's the header) and remove it.
                if (s.Blocks.Count > 0)
                    s.Blocks.RemoveAt(0);
                // Insert the new header into the each section.
                s.HeadersFooters.Add(header.Clone(true));                
            }

            dc.Save(outFile);

            // Open the results for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(inpFile) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });

        }
    }
}