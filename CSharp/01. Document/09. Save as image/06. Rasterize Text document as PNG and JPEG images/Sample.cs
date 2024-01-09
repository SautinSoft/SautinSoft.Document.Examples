using System;
using System.IO;
using SautinSoft.Document;
using SkiaSharp;

namespace Example
{
    class Program
    {      
        static void Main(string[] args)
        {
            // Get your free 30-day key here:   
            // https://sautinsoft.com/start-for-free/

            RasterizeTxtToPicture();         
        }

        /// <summary>
        /// Rasterizing - save Text document as PNG and JPEG images.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/rasterize-save-text-document-as-png-jpeg-images-net-csharp-vb.php
        /// </remarks>
        static void RasterizeTxtToPicture()
        {
            // In this example we'll how rasterize/save 1st and 2nd pages of Text document
            // as PNG and JPEG images.
            string inputFile = @"..\..\..\example.txt";
            string jpegFile = "Result.jpg";
            string pngFile = "Result.png";
            // The file format is detected automatically from the file extension: ".txt".
            // But as shown in the example below, we can specify TxtLoadOptions as 2nd parameter
            // to explicitly set that a loadable document has Text format.
            DocumentCore dc = DocumentCore.Load(inputFile, new TxtLoadOptions() { Encoding = System.Text.Encoding.UTF8 });
            
            DocumentPaginator documentPaginator = dc.GetPaginator(new PaginatorOptions() { UpdateFields = true });

            int dpi = 300;

            int pagesToRasterize = 2;
            int currentPage = 1;
            
            foreach (DocumentPage page in documentPaginator.Pages)
            {
                // Save the page into Bitmap image with specified dpi and background.
                var DPI = new ImageSaveOptions();
                DPI.DpiX = 300;
                DPI.DpiY = 300;
                // Save the 1st document page to the file in PNG format.
                SkiaSharp.SKBitmap picture = page.Rasterize(DPI, SautinSoft.Document.Color.White);

                // Save the Bitmap to a PNG file.
                if (currentPage == 1)
                    picture.Encode(new FileStream(pngFile, FileMode.Create), SkiaSharp.SKEncodedImageFormat.Png, 100);
                else if (currentPage == 2)
                    // Save the Bitmap to a JPEG file.
                    picture.Encode(new FileStream(jpegFile, FileMode.Create), SkiaSharp.SKEncodedImageFormat.Jpeg, 100);

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