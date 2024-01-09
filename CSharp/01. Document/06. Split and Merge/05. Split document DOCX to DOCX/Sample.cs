using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free 30-day key here:   
            // https://sautinsoft.com/start-for-free/

            SplitDocumentByPages();
        }
        /// <summary>
        /// Loads a document and split it by separate pages. Saves the each page into DOCX format.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/split-docx-document-by-pages-in-docx-format-net-csharp-vb.php
        /// </remarks>
        static void SplitDocumentByPages()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            string folderPath = Path.GetFullPath(@"Result-files");
            DocumentPaginator dp = dc.GetPaginator();
            for (int i = 0; i < dp.Pages.Count; i++)
            {
                DocumentPage page = dp.Pages[i];
                Directory.CreateDirectory(folderPath);

                // Save the each page to DOCX format.
                page.Save(folderPath + @"\Page - " + i.ToString() + ".docx", SaveOptions.DocxDefault);
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(folderPath) { UseShellExecute = true });
        }
    }
}