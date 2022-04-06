using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            SaveToImage();
        }
        /// <summary>
        /// Loads a document and saves all pages as images.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/save-document-as-image-net-csharp-vb.php
        /// </remarks>
        static void SaveToImage()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            string folderPath = Path.GetFullPath(@"Result-files");

            DocumentPaginator dp =  dc.GetPaginator();

            for (int i = 0; i < dp.Pages.Count; i++)
            {
                DocumentPage page = dp.Pages[i];
                // For example, set DPI: 72, Background: White.
                Bitmap image = page.Rasterize(72, SautinSoft.Document.Color.White);
                Directory.CreateDirectory(folderPath);
                image.Save(folderPath+@"\Page - "+i.ToString()+".png", ImageFormat.Png);
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(folderPath) { UseShellExecute = true });
        }
    }
}