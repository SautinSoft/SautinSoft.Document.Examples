using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free 100-day key here:   
            // https://sautinsoft.com/start-for-free/

            SeparateDocumentToHtmlPages();
        }
        /// <summary>
        /// Load a document and save all pages as separate HTML documents.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/pagination-save-document-pages-as-html-net-csharp-vb.php
        /// </remarks>
        static void SeparateDocumentToHtmlPages()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            string folderPath = Path.GetFullPath(@"Result-files");
            DocumentPaginator dp = dc.GetPaginator();
            for (int i = 0; i < dp.Pages.Count; i++)
            {
                DocumentPage page = dp.Pages[i];
                Directory.CreateDirectory(folderPath);

                // Save the each page into HTML format.
                page.Save(folderPath + @"\Page - " + (i + 1).ToString() + ".html", SaveOptions.HtmlFixedDefault);
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(folderPath) { UseShellExecute = true });
        }
    }
}