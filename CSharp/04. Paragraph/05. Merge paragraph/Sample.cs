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
            MergeParagraphs();
        }
        /// <summary>
        /// Merge all paragraphs into a single in an existing PDF document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/merge-paragraphs-in-pdf-document-net-csharp-vb.php
        /// </remarks>
        static void MergeParagraphs()
        {
            string inpFile = @"..\..\..\example.pdf";
            string outFile = @"Result.pdf";
            DocumentCore dc = DocumentCore.Load(inpFile);

            Paragraph firstPar = dc.GetChildElements(true, ElementType.Paragraph).First() as Paragraph;

            int lastIndex = firstPar.Inlines.Count;

            foreach (Paragraph par in dc.GetChildElements(true, ElementType.Paragraph).Reverse().Where(p => p != firstPar))
            {
                int last = lastIndex;
                foreach(Inline inline in par.Inlines)
                {
                    firstPar.Inlines.Insert(last++, inline.Clone(true));
                }
                par.Content.Delete();
            }

            dc.Save(outFile);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(inpFile) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}