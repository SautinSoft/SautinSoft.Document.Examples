using System;
using System.IO;
using SautinSoft.Document;
using System.Drawing;

namespace Example
{
    class Program
    {      
        static void Main(string[] args)
        {
            RasterizeDocxToPicture();         
        }

        /// <summary>
        /// Rasterizing - save DOCX document as PNG and JPEG images.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/rasterize-save-docx-document-as-png-jpeg-images-net-csharp-vb.php
        /// </remarks>
        static void RasterizeDocxToPicture()
        {
            // In this example we'll how rasterize/save 1st and 2nd pages of DOCX document
            // as PNG and JPEG images.
            string inputFile = @"..\..\..\example.docx";
            string jpegFile = "Result.jpg";
            string pngFile = "Result.png";
            // The file format is detected automatically from the file extension: ".docx".
            // But as shown in the example below, we can specify DocxLoadOptions as 2nd parameter
            // to explicitly set that a loadable document has Docx format.
            DocumentCore dc = DocumentCore.Load(inputFile);
            
            DocumentPaginator documentPaginator = dc.GetPaginator(new PaginatorOptions() { UpdateFields = true });

            int dpi = 300;

            int pagesToRasterize = 2;
            int currentPage = 1;
            
            foreach (DocumentPage page in documentPaginator.Pages)
            {
                // Save the page into Bitmap image with specified dpi and background.
                Bitmap picture = page.Rasterize(dpi, SautinSoft.Document.Color.White);

                // Save the Bitmap to a PNG file.
                if (currentPage == 1)
                    picture.Save(pngFile);
                else if (currentPage == 2)
                // Save the Bitmap to a JPEG file.
                    picture.Save(jpegFile, System.Drawing.Imaging.ImageFormat.Jpeg);

                currentPage++;

                if (currentPage > pagesToRasterize)
                    break;
            }

            // Open the results for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pngFile) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(jpegFile) { UseShellExecute = true });
        }
    }
}