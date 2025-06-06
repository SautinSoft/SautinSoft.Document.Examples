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
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            ContentRangeManipulation();
        }
        /// <summary>
        /// Adds two paragraphs by different ways: using ContentRange and as Element.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/contentrange-manipulation.php
        /// </remarks>
        static void ContentRangeManipulation()
        {
            string filePath = @"Result.docx";

            // Let's create a new document.
            DocumentCore dc = new DocumentCore();

            // Way 1: Add new paragraph using the property Content (class ContentRange).
            Paragraph par = new Paragraph(dc, "This is paragraph. ");
            par.ParagraphFormat.BackgroundColor = Color.Yellow;
            // Note, our Paragraph will be cloned and the clone will be inserted. 
            // The property Content (class ContentRange) each time clones the inserting object.
            dc.Content.End.Insert(par.Content);

            // Way 2: Add paragraph into the Block collection as Element.
            // Note, you can't insert the one Element two times.
            // The cloning does not occur here.
            par.ParagraphFormat.BackgroundColor = Color.Red;
            dc.Sections[0].Blocks.Add(par);

            // Again Way 1: Change Background color to blue and insert 
            // the copy (clone) of our paragraph using Content (class ContentRange).
            par.ParagraphFormat.BackgroundColor = Color.Blue;
            dc.Content.End.Insert(par.Content);

            // Change background to 'Green' for the paragraph.
            // This affect only to the paragraph inserted by 'Way 2' - as Element.
            // Because in 'Way 1' we added clones using ContentRange.
            par.ParagraphFormat.BackgroundColor = Color.Green;

            // Save our document.
            dc.Save(filePath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}