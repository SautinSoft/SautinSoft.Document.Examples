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
            FindParagraph();
        }
        /// <summary>
        /// Find all paragraphs aligned by center in DOCX document and mark it by yellow.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/find-paragraphs-in-docx-document-net-csharp-vb.php
        /// </remarks>
        static void FindParagraph()
        {
            string filePath = @"..\..\..\example.docx";
            string fileResult = @"Result.docx";
            DocumentCore dc = DocumentCore.Load(filePath);

            foreach (Paragraph par in dc.GetChildElements(true, ElementType.Paragraph).
                Where(p => (p as Paragraph).ParagraphFormat.Alignment == HorizontalAlignment.Center))
            {
                par.ParagraphFormat.BackgroundColor = Color.Yellow;
            }
            dc.Save(fileResult);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fileResult) { UseShellExecute = true });

        }
    }
}