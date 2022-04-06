using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            SplitDocumentByPages();
        }
        /// <summary>
        /// Loads a document and split it by separate pages. Saves the each page into PDF format.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/split-docx-document-by-pages-in-pdf-format-net-csharp-vb.php
        /// </remarks>
        static void SplitDocumentByPages()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            string folderPath = Path.GetFullPath(@"Result-files");
            DocumentPaginator dp =  dc.GetPaginator();
            for (int i = 0; i < dp.Pages.Count; i++)
            {
                DocumentPage page = dp.Pages[i];
                Directory.CreateDirectory(folderPath);

                // Save the each page to PDF format.
                page.Save(folderPath + @"\Page - " + i.ToString() + ".pdf", SaveOptions.PdfDefault);
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(folderPath) { UseShellExecute = true });
        }
    }
}