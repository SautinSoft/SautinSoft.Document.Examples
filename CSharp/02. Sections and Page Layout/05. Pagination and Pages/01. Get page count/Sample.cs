using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            Pagination();
        }
        /// <summary>
        /// Loads a document and applies Pagination to get separate pages.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/pagination.php
        /// </remarks>
        static void Pagination()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            string folderPath = Path.GetFullPath(@"Result-files");
            DocumentPaginator dp =  dc.GetPaginator();
            for (int i = 0; i < dp.Pages.Count; i++)
            {
                DocumentPage page = dp.Pages[i];
                Directory.CreateDirectory(folderPath);

                // Save the each page into PDF format.
                page.Save(folderPath + @"\Page - " + i.ToString() + ".pdf", SaveOptions.PdfDefault);
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(folderPath) { UseShellExecute = true });
        }
    }
}