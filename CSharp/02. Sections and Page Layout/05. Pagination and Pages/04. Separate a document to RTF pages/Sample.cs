using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            SeparateDocumentToRtfPages();
        }
        /// <summary>
        /// Load a document and save all pages as separate RTF documents.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/pagination-save-document-pages-as-rtf-net-csharp-vb.php
        /// </remarks>
        static void SeparateDocumentToRtfPages()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            string folderPath = Path.GetFullPath(@"Result-files");
            DocumentPaginator dp = dc.GetPaginator();
            for (int i = 0; i < dp.Pages.Count; i++)
            {
                DocumentPage page = dp.Pages[i];
                Directory.CreateDirectory(folderPath);

                // Save the each page into RTF format.
                page.Save(folderPath + @"\Page - " + (i + 1).ToString() + ".rtf", SaveOptions.RtfDefault);
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(folderPath) { UseShellExecute = true });
        }
    }
}