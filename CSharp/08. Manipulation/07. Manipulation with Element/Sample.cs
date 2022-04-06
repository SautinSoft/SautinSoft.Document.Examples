using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using SautinSoft.Document.Tables;
using System.Linq;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            ElementManipulation();
        }
        /// <summary>
        /// Create a document and add paragraphs as content and as element.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/element-manipulation.php
        /// </remarks>
        static void ElementManipulation()
        {
            string filePath = @"Result.docx";

            // Let's create a new document.
            DocumentCore dc = new DocumentCore();
            Paragraph par = new Paragraph(dc, "This is the first paragraph.");

            // Insert the clone of our Paragraph using ContentRange.
            dc.Content.End.Insert(par.Content);

            // Add our Paragraph in Block collection as Element.
            dc.Sections[0].Blocks.Add(par);

            // Again, insert the clone of our Paragraph using ContentRange.
            dc.Content.End.Insert(par.Content);

            // Change text in our Paragraph
            (par.Inlines[0] as Run).Text = "Now we are in the second paragraph.";

            // Find 3rd paragraph and change text in it.
            ((par.NextSibling as Paragraph).Inlines[0] as Run).Text = "This is the third paragraph.";

            // Save our document.
            dc.Save(filePath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}