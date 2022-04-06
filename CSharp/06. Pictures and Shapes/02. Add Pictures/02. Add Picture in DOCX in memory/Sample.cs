using System;
using System.IO;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            AddPictureToDocxInMemory();
        }

        /// <summary>
        /// How to add picture into an existing DOCX document using MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/add-picture-in-docx-document-in-memory-net-csharp-vb.php
        /// </remarks>
        public static void AddPictureToDocxInMemory()
        {
            // We're using files here only to retrieve the data from them and show the results.
            // The whole process will be done completely in memory using MemoryStream.
            string inputFile = @"..\..\..\example.docx";
            string outputDocxFile = @"Result.docx";
            string outputPdfFile = @"Result.pdf";
            string pictPath = @"..\..\..\picture.jpg";

            // 1. Load the input data into memory (from DOCX document and the picture).
            byte[] inputDocxBytes = File.ReadAllBytes(inputFile);
            byte[] pictBytes = File.ReadAllBytes(pictPath);

            // 2. Create new MemoryStream with DOCX and load it into DocumentCore.
            DocumentCore dc = null;
            using (MemoryStream msDocx = new MemoryStream(inputDocxBytes))
            {
                dc = DocumentCore.Load(msDocx, new DocxLoadOptions());
            }

            // 3. Create new Memory Stream, and Picture object for the picture.
            Picture pict = null;

            // Set the picture size in mm.
            int width = 40;
            int height = 40;
            Size size = new Size(LengthUnitConverter.Convert(width, LengthUnit.Millimeter, LengthUnit.Point),
                         LengthUnitConverter.Convert(height, LengthUnit.Millimeter, LengthUnit.Point));

            // Set the picture layout from the (left, top) page corner.
            int fromLeftMm = 140;
            int fromTopMm = 180;

            // Floating layout means that the Picture (or Shape) is positioned by coordinates.
            FloatingLayout fl = new FloatingLayout(new HorizontalPosition(fromLeftMm, LengthUnit.Millimeter, HorizontalPositionAnchor.Page),
                new VerticalPosition(fromTopMm, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), size);

            // Load the picture
            using (MemoryStream msPict = new MemoryStream(pictBytes))
            {
                pict = new Picture(dc, fl, msPict, PictureFormat.Jpeg);
            }

            // Set the wrapping style.
            (pict.Layout as FloatingLayout).WrappingStyle = WrappingStyle.Tight;

            // Add our picture into the 1st section.
            Section sect = dc.Sections[0];
            sect.Content.End.Insert(pict.Content);

            // Save our document into DOCX format using MemoryStream.
            using (MemoryStream msDocxResult = new MemoryStream())
            {
                dc.Save(msDocxResult, new DocxSaveOptions());

                // To show the result save our msDocxResult into file.
                File.WriteAllBytes(outputDocxFile, msDocxResult.ToArray());

                // Open the result for demonstration purposes.
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputDocxFile) { UseShellExecute = true });
            }

            // Save our document into PDF format using MemoryStream.
            using (MemoryStream msPdfResult = new MemoryStream())
            {
                dc.Save(msPdfResult, new PdfSaveOptions());

                // To show the result save our msPdfResult into file.
                File.WriteAllBytes(outputPdfFile, msPdfResult.ToArray());

                // Open the result for demonstration purposes.
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPdfFile) { UseShellExecute = true });
            }
        }
    }
}