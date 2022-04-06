using System;
using System.IO;
using System.Linq;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            FindAndReplaceInParagraphs();
        }
        /// <summary>
        /// Find and replace a specific text in all paragraphs in PDF document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/replace-text-paragraphs-in-pdf-document-net-csharp-vb.php
        /// </remarks>
        static void FindAndReplaceInParagraphs()
        {
            string filePath = @"..\..\..\example.pdf";
            string fileResult = @"Result.pdf";
            DocumentCore dc = DocumentCore.Load(filePath);
            foreach (Paragraph par in dc.GetChildElements(true, ElementType.Paragraph))
                foreach (ContentRange item in par.Content.Find("old text").Reverse())
                {
                    item.Replace("new text");
                }
            dc.Save(fileResult);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fileResult) { UseShellExecute = true });
        }
    }
}