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
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            RasterizeRtfToPicture();
        }

        /// <summary>
        /// Rasterizing - save RTF document as PNG and JPEG images.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/rasterize-save-rtf-document-as-png-jpeg-images-net-csharp-vb.php
        /// </remarks>
        static void RasterizeRtfToPicture()
        {
            // In this example we'll how rasterize/save 1st and 2nd pages of RTF document
            // as PNG and JPEG images.
            string inputFile = @"..\..\..\example.rtf";
            string jpegFile = "Result.jpeg";
            string pngFile = "Result.png";
            // The file format is detected automatically from the file extension: ".rtf".
            // But as shown in the example below, we can specify RtfLoadOptions as 2nd parameter
            // to explicitly set that a loadable document has Rtf format.
            DocumentCore dc = DocumentCore.Load(inputFile, new RtfLoadOptions());

            DocumentPaginator documentPaginator = dc.GetPaginator(new PaginatorOptions() { UpdateFields = true });

            int pagesToRasterize = 2;
            int currentPage = 1;

            foreach (DocumentPage page in documentPaginator.Pages)
            {
                // Save to a PNG file.
                if (currentPage == 1)
                    page.Save(pngFile, new ImageSaveOptions
                    {
                        DpiX = 72,
                        DpiY = 72,
                        Format = ImageSaveFormat.Png
                    });
                // Save to a JPEG file.
                else if (currentPage == 2)
                    page.Save(jpegFile, new ImageSaveOptions
                    {
                        DpiX = 72,
                        DpiY = 72,
                        Format = ImageSaveFormat.Jpeg
                    });

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