using System.Linq;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            ReplaceText();
        }
        /// <summary>
        /// Replace a specific text in an existing DOCX document.
        /// </summary>
        /// </remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/replace-text-in-docx-document-net-csharp-vb.php
        /// </remarks>
        static void ReplaceText()
        {
            string filePath = @"..\..\..\example.docx";
            string fileResult = @"Result.docx";
            string searchText = "document";
            string replaceText = "book";
            DocumentCore dc = DocumentCore.Load(filePath);
            foreach (ContentRange cr in dc.Content.Find(searchText).Reverse())
            {
                // Replace "document" to "book";
                // Mark "book" by yellow.
                cr.Replace(replaceText, new CharacterFormat() { BackgroundColor = Color.Yellow});
            }
            dc.Save(fileResult);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fileResult) { UseShellExecute = true });

        }
    }
}