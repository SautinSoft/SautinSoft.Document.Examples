using System.IO;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            AddHeaderFooter();
        }

        /// <summary>
        /// How to add a header and footer into PDF document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/add-header-and-footer-in-pdf-net-csharp-vb.php
        /// </remarks>
        static void AddHeaderFooter()
        {
            string inpFile = @"..\..\..\shrek.pdf";
            string outFile = "Shrek with header and footer.pdf";
            DocumentCore dc = DocumentCore.Load(inpFile);

            // Create new header with formatted text.
            HeaderFooter header = new HeaderFooter(dc, HeaderFooterType.HeaderDefault);
            header.Content.Start.Insert("Shrek and Donkey", new CharacterFormat() { Size = 14.0, FontColor = Color.Brown });
            foreach (Section s in dc.Sections)
            {
                s.HeadersFooters.Add(header.Clone(true));
            }

            // Create new footer with formatted text.
            HeaderFooter footer = new HeaderFooter(dc, HeaderFooterType.FooterDefault);
            footer.Content.Start.Insert("Fiona.", new CharacterFormat() { Size = 14.0, FontColor = Color.Blue });
            foreach (Section s in dc.Sections)
            {
                s.HeadersFooters.Add(footer.Clone(true));
            }

            dc.Save(outFile);

            // Open the PDF documents for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(inpFile) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}