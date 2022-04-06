using System.Linq;
using System.IO;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertPictureToCustomPages();
        }
        /// <summary>
        /// Insert a picture to custom pages into existing DOCX document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-picture-jpg-image-to-custom-docx-page-net-csharp-vb.php
        /// </remarks>
        static void InsertPictureToCustomPages()
        {
            // In this example we'll insert the picture to 1st and 3rd pages
            // of DOCX document into specific positions.

            string inpFile = @"..\..\..\example.docx";
            string outFile = @"Result.docx";
            string pictFile = @"..\..\..\picture.jpg";

            DocumentCore dc = DocumentCore.Load(inpFile);
            DocumentPaginator dp = dc.GetPaginator();


            // Step 1: Put the picture to 1st page.

            // Create the Picture object from Jpeg file.
            Picture pict = new Picture(dc, pictFile);

            // Specify the picture size and position.
            pict.Layout = FloatingLayout.Floating(
                new HorizontalPosition(70, LengthUnit.Millimeter, HorizontalPositionAnchor.Margin),
                new VerticalPosition(23, LengthUnit.Millimeter, VerticalPositionAnchor.Margin),
                new Size(LengthUnitConverter.Convert(1, LengthUnit.Inch, LengthUnit.Point),
                          pict.Layout.Size.Height * LengthUnitConverter.Convert(1, LengthUnit.Inch, LengthUnit.Point) / pict.Layout.Size.Width));

            // Put the picture behind the text
            (pict.Layout as FloatingLayout).WrappingStyle = WrappingStyle.BehindText;

            // Find the 1st Element in the 1st page.
            Element e1 = dp.Pages[0].GetElementFrames().FirstOrDefault(e => e.Element is Run).Element;

            // Insert the picture at this Element.
            e1.Content.End.Insert(pict.Content);


            // Step 2: Put the picture to 3rd page.
            if (dp.Pages.Count >= 3)
            {
                // Find the 1st Element on the 3rd page.
                Element e2 = dp.Pages[2].GetElementFrames().FirstOrDefault(e => e.Element is Run).Element;

                // Create another picture
                Picture pict2 = new Picture(dc, pictFile);
                pict2.Layout = FloatingLayout.Floating(
                    new HorizontalPosition(10, LengthUnit.Millimeter, HorizontalPositionAnchor.Margin),
                    new VerticalPosition(20, LengthUnit.Millimeter, VerticalPositionAnchor.Margin),
                    new Size(LengthUnitConverter.Convert(1, LengthUnit.Inch, LengthUnit.Point),
                             pict2.Layout.Size.Height * LengthUnitConverter.Convert(1, LengthUnit.Inch, LengthUnit.Point) / pict2.Layout.Size.Width)
                             );
                (pict2.Layout as FloatingLayout).WrappingStyle = WrappingStyle.BehindText;

                // Insert the picture at this Element.
                e2.Content.End.Insert(pict2.Content);
            }
            // Save the document as new DOCX and open it.
            dc.Save(outFile);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}