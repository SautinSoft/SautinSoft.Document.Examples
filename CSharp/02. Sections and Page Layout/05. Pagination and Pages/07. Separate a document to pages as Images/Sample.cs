using System.IO;
using SautinSoft.Document;
using SkiaSharp;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free 100-day key here:   
            // https://sautinsoft.com/start-for-free/

            SeparateDocumentToImagePages();
        }
        /// <summary>
        /// Load a document and save all pages as separate PNG &amp; Jpeg images.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/pagination-save-document-pages-as-png-jpg-jpeg-net-csharp-vb.php
        /// </remarks>
        static void SeparateDocumentToImagePages()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            string folderPath = Path.GetFullPath(@"Result-files");
            DocumentPaginator dp = dc.GetPaginator();
            for (int i = 0; i < dp.Pages.Count; i++)
            {
                DocumentPage page = dp.Pages[i];
                Directory.CreateDirectory(folderPath);

                // Save the each page as Bitmap.
                ImageSaveOptions op = new ImageSaveOptions();
                op.DpiX = 300;
                op.DpiY = 300;
                SKBitmap bmp = page.Rasterize(op, SautinSoft.Document.Color.White);
                // Save the bitmap to PNG and JPEG.
                bmp.Encode(new FileStream(folderPath + @"\Page (PNG) - " + (i + 1).ToString() + ".png", FileMode.Create), SKEncodedImageFormat.Png, 100);
                bmp.Encode(new FileStream(folderPath + @"\Page (Jpeg) - " + (i + 1).ToString() + ".jpg", FileMode.Create), SKEncodedImageFormat.Jpeg, 100);
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(folderPath) { UseShellExecute = true });
        }
    }
}