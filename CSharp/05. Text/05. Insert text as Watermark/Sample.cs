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
            InsertTextAsWatermark();
        }
        /// <summary>
        /// How to insert a text watermark in the existing PDF, DOCX, any document to the each page.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-text-watermark-in-the-existing-pdf-docx-document-to-each-page-net-csharp-vb.php
        /// </remarks>
        static void InsertTextAsWatermark()
        {
            // Let us say, we want to insert a textual "Watermark" at the each page of the
            // DOCX document and specify angle of 45 degree for it.
            
            // If we'll insert the text at the document header (or footer), so it will appear on the each page.
            
            // Also let's insert our Watermark behind the main content.
            string inpFile = @"..\..\..\example.docx";
            string outFile = @"Result.docx";

            // 1. Load an existing DOCX document.
            DocumentCore dc = DocumentCore.Load(inpFile);

            // 2. Create a Shape with our Watermark text.
            // Place the watermark:
            // 30mm - from the page left;            
            // 150mm - from the page top.
            // 60 - angle.
            float posFromLeft = 30f;
            float posFromTop = 150f;
            float angle = -60f;
            // Size of the Shape, 200mm x 40mm.            
            SautinSoft.Document.Drawing.Size size = new Size(200f, 40f, LengthUnit.Millimeter);
            Shape watermark = new Shape(dc, new FloatingLayout(new HorizontalPosition(posFromLeft, LengthUnit.Millimeter, HorizontalPositionAnchor.Page),
            new VerticalPosition(posFromTop, LengthUnit.Millimeter, VerticalPositionAnchor.Page), size));
            // Rotate shape.
            watermark.Rotation = angle;
            // Create the text.
            Run text = new Run(dc, "Watermark!", new CharacterFormat()
            {
                Size = 100f,
                FontColor = Color.Black,
                FontName = "Arial"
            });
            watermark.Text.Blocks.Add(new Paragraph(dc, text));
            // Set shape Behind the main document contents.
            (watermark.Layout as FloatingLayout).WrappingStyle = WrappingStyle.BehindText;
            // Remove the shape borders.
            watermark.Outline.Fill.SetEmpty();

            // 3. Iterate through Sections, and insert our Watermark to the default header of the each section.
            foreach (Section section in dc.Sections)
            {
                // 2.1. Check the document header, maybe is it already exist?
                var header = dc.Sections[0].HeadersFooters[HeaderFooterType.HeaderDefault];

                if (header == null)
                {
                    // Create a new header, add it into the section.
                    header = new HeaderFooter(dc, HeaderFooterType.HeaderDefault);
                    // Add the header to the section.
                    section.HeadersFooters.Add(header);
                }
                // Add the watermark to the header.
                header.Content.End.Insert(watermark.Content);
            }

            // 4. Save the document back.
            dc.Save(outFile);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}