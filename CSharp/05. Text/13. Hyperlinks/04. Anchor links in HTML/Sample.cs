using System.Text;
using System.Linq;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            AnchorLinks();
        }

        /// <summary>
        /// Insert anchor links inside the HTML page using C# and .NET
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-anchor-links-inside-an-html-page-using-csharp-vb-net.php
        /// </remarks>		
        public static void AnchorLinks()
        {
            // P.S. If you are using MS Word, to display bookmarks:
            // File -> Options -> Advanced -> On the "Show document content" check "Show bookmarks".
            string documentPath = @"AnchorLinks.html";

            // Let's create a new document.
            DocumentCore dc = new DocumentCore();
            dc.Sections.Add(new Section(dc));

            // Add an anchor links to the end of the document.
            dc.Sections[0].Blocks.Add(
            new Paragraph(dc,
            new Hyperlink(dc, "IdEnd", "Document End") { IsBookmarkLink = true }));

            // Add 100 paragraphs
            for (int i = 0; i < 100; i++)            
                dc.Sections[0].Blocks.Add(new Paragraph(dc, new Run(dc, $"Paragraph {i + 1}")));

            dc.Sections[0].Blocks.Add(
                new Paragraph(dc,
                    new BookmarkStart(dc, "IdEnd"),
                    new Run(dc, "The document end."),
                    new BookmarkEnd(dc, "IdEnd"))); 

            // Let's save the document as HTML.
            dc.Save(documentPath, new HtmlFlowingSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
        
    }
}