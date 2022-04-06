using System;
using System.IO;
using System.Linq;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertTextByCoordinates();
        }
        /// <summary>
        /// How to insert a text in the existing PDF, DOCX, any document by specific (x,y) coordinates
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-text-in-the-existing-pdf-docx-document-by-specific-x-y-coordinates-net-csharp-vb.php
        /// </remarks>
        static void InsertTextByCoordinates()
        {
            // Let us say, we want to insert the text "Hello World!" into:
            // the pages 2,3;
            // 50mm - from the left;
            // 30mm - from the top.
            // Also the text must be inserted Behind the existing text.
            string inpFile = @"..\..\..\example.docx";
            string outFile = @"Result.docx";

            // 1. Load an existing document
            DocumentCore dc = DocumentCore.Load(inpFile);

            // 2. Get document pages
            var paginator = dc.GetPaginator(new PaginatorOptions() { UpdateFields = true });
            var pages = paginator.Pages;

            // 3. Check that we at least 3 pages in our document.
            if (pages.Count < 3)
            {
                Console.WriteLine("The document contains less than 3 pages!");
                Console.ReadKey();
                return;
            }
            // 50mm - from the left;            
            // 30mm - from the top.
            float posFromLeft = 50f;
            float posFromTop = 30f;

            // Insert the text "Hello World!" into the page 2.
            Run text1 = new Run(dc, "Hello World!", new CharacterFormat() { Size = 36, FontColor = Color.Red, FontName = "Arial" });
            InsertShape(dc, pages[1], text1, posFromLeft, posFromTop);

            // Insert the text "Hej Världen!" into the page 3.
            Run text2 = new Run(dc, "Hej Världen!", new CharacterFormat() { Size = 36, FontColor = Color.Orange, FontName = "Arial" });
            InsertShape(dc, pages[2], text2, posFromLeft, posFromTop);

            // 4. Save the document back.
            dc.Save(outFile);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
        /// <summary>
        /// Inserts a Shape with text into the specific page using coordinates.
        /// </summary>
        /// <param name="dc">The document.</param>
        /// <param name="page">The specific page to insert the Shape.</param>
        /// <param name="text">The formatted text (Run object).</param>
        /// <param name="posFromLeftMm">The distance in mm from left corner of the page</param>
        /// <param name="posFromTopMm">The distance in mm from top corner of the page</param>
        static void InsertShape(DocumentCore dc, DocumentPage page, Run text, float posFromLeftMm, float posFromTopMm)
        {
            HorizontalPosition hp = new HorizontalPosition(posFromLeftMm, LengthUnit.Millimeter, HorizontalPositionAnchor.Page);
            VerticalPosition vp = new VerticalPosition(posFromTopMm, LengthUnit.Millimeter, VerticalPositionAnchor.Page);
            // 100 x 30 mm
            float shapeWidth = 100f;
            float shapeHeight = 30f;
            SautinSoft.Document.Drawing.Size size = new Size(shapeWidth, shapeHeight, LengthUnit.Millimeter);
            Shape shape = new Shape(dc, new FloatingLayout(hp, vp, size));
            shape.Text.Blocks.Add(new Paragraph(dc, text));
            // Set shape Behind the text.
            (shape.Layout as FloatingLayout).WrappingStyle = WrappingStyle.BehindText;
            // Remove the shape borders.
            shape.Outline.Fill.SetEmpty();
            // Insert shape into the page.
            page.Content.End.Insert(shape.Content);
        }
    }
}