using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            ChangeHeader();
        }
        /// <summary>
        /// How to edit Header and Footer in PDF file
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/edit-header-and-footer-in-pdf-net-csharp-vb.php
        /// </remarks>
        static void ChangeHeader()
        {
            string inpFile = @"..\..\somebody.pdf";
            string outFile = "With changed header.pdf";
            DocumentCore dc = DocumentCore.Load(inpFile);

            // Create new header with formatted text.
            HeaderFooter header = new HeaderFooter(dc, HeaderFooterType.HeaderDefault);
            header.Content.Start.Insert("Modified : 1 April 2020", new CharacterFormat() { Size = 14.0, FontColor = Color.DarkGreen });
            foreach (Section s in dc.Sections)
            {
                if (s.Blocks.Count > 0)
                    s.Blocks.RemoveAt(1);
                s.HeadersFooters.Add(header.Clone(true));
            }
            dc.Save(outFile);

            // Open the results for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(inpFile) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });

        }
    }
}