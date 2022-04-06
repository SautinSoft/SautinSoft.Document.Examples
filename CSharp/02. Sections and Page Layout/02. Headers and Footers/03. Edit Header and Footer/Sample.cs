using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            ChangeHeaderAndFooter();
        }
        /// <summary>
        /// How to edit Header and Footer in PDF file
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/edit-header-and-footer-in-pdf-net-csharp-vb.php
        /// </remarks>
        static void ChangeHeaderAndFooter()
        {
            string inpFile = @"..\..\..\somebody.pdf";
            string outFile = "With changed header and footer.pdf";
            DocumentCore dc = DocumentCore.Load(inpFile);

            // Create new header with formatted text.
            HeaderFooter header = new HeaderFooter(dc, HeaderFooterType.HeaderDefault);
            header.Content.Start.Insert("Modified: 1 April 2020", new CharacterFormat() { Size = 14.0, FontColor = Color.DarkGreen });

            // Create the footer with orange text, with font name Elephant and size of 50 pt.
            HeaderFooter footer = new HeaderFooter(dc, HeaderFooterType.FooterDefault);
            Paragraph p = new Paragraph(dc, new Run(dc, "Last modified: 1st June 2021",
                new CharacterFormat() { Size = 50.0, FontColor = Color.Orange, FontName = "Elephant" }));
            p.ParagraphFormat.Alignment = HorizontalAlignment.Left;
            footer.Blocks.Add(p);

            foreach (Section s in dc.Sections)
            {
                if (s.Blocks.Count > 0)
                    s.Blocks.RemoveAt(1);
                s.HeadersFooters.Add(header.Clone(true));
                s.HeadersFooters.Add(footer.Clone(true));
            }
            dc.Save(outFile);

            // Open the results for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(inpFile) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });

        }
    }
}