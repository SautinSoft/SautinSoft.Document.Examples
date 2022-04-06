using System.IO;
using System.Collections.Generic;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertImagesOnEachPage();
        }
        /// <summary>
        /// Insert images on each page of the PDF file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-insert-images-on-each-document-page-csharp-vb-net.php
        /// </remarks>
        static void InsertImagesOnEachPage()
        {
            string inpfFile = @"..\..\..\example.pdf";
            string pctFile = @"..\..\..\signature.png";
            string outFile = @"Result.pdf";

            // This example is acceptable for PDF documents.
            // Because when we're loading PDF documents using DocumentCore the each PDF-page
            // we'll be stored in own Section object.
            // In other words, the each Section represents the separate PDF-page.
            DocumentCore dc = DocumentCore.Load(inpfFile);

            // Load the Picture from a file.
            Picture pict = new Picture(dc, pctFile);

            // In this example we'll place the image in three (3) 
            // different places to the each document page.

            // Let's create three layouts
            List<Layout> layouts = new List<Layout>()
            {
                // Layout 1.
                // Horizontal: 10mm from page left.
                // Vertical: 260mm from top margin.
                // Size: 2cm * 1 cm.
                FloatingLayout.Floating(
                new HorizontalPosition(10, LengthUnit.Millimeter, HorizontalPositionAnchor.Page),
                new VerticalPosition(260, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                new Size(LengthUnitConverter.Convert(2, LengthUnit.Centimeter, LengthUnit.Point),
                         LengthUnitConverter.Convert(1, LengthUnit.Centimeter, LengthUnit.Point))
                         ),
                // Layout 2.
                // Horizontal: 180mm from page left.
                // Vertical: 10mm from top margin.
                // Size: 3cm * 2cm.
                FloatingLayout.Floating(
                new HorizontalPosition(180, LengthUnit.Millimeter, HorizontalPositionAnchor.Page),
                new VerticalPosition(10, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                new Size(LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.Point),
                         LengthUnitConverter.Convert(2, LengthUnit.Centimeter, LengthUnit.Point))
                         ),

                // Layout 3.
                // Horizontal: 150mm from page left.
                // Vertical: 150mm from page top.
                // Size: 3cm * 3cm.
                FloatingLayout.Floating(
                new HorizontalPosition(150, LengthUnit.Millimeter, HorizontalPositionAnchor.Page),
                new VerticalPosition(150, LengthUnit.Millimeter, VerticalPositionAnchor.Page),
                new Size(LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.Point),
                         LengthUnitConverter.Convert(3, LengthUnit.Centimeter, LengthUnit.Point))
                         )};

            // Iterate by Sections (PDF pages in our case).
            foreach (Section s in dc.Sections)
            {
                // Insert our pictures in different places.
                foreach (FloatingLayout fl in layouts)
                {
                    pict.Layout = new FloatingLayout(fl.HorizontalPosition, fl.VerticalPosition, fl.Size);
                    // Place the picture behind the text.
                    (pict.Layout as FloatingLayout).WrappingStyle = WrappingStyle.BehindText;

                    // Here we insert the Picture content at the 1st Block element (Paragraph or Table).
                    s.Blocks[0].Content.Start.Insert(pict.Content);
                }
            }

            dc.Save(outFile, new PdfSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}
