using System.IO;
using SautinSoft.Document;
using System.Drawing;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
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
                Bitmap bmp = page.Rasterize(300, SautinSoft.Document.Color.White);
                // Save the bitmap to PNG and JPEG.
                bmp.Save(folderPath + @"\Page (PNG) - " + (i + 1).ToString() + ".png");
                bmp.Save(folderPath + @"\Page (Jpeg) - " + (i + 1).ToString() + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(folderPath) { UseShellExecute = true });
        }
    }
}