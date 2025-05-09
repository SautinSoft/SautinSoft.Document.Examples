using System;
using System.IO;
using System.Runtime;
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

                // To get high-quality image, lets set 72 dpi.
                var DPI = new ImageSaveOptions();
                DPI.DpiX = 72;
                DPI.DpiY = 72;
                
                // Convert the page into PNG image.
                Directory.CreateDirectory(folderPath);
                page.Save(folderPath + @"\Page - " + i.ToString() + ".png", DPI);

            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(folderPath) { UseShellExecute = true });
        }
    }
}