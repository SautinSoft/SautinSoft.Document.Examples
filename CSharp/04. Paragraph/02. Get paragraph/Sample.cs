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
            GetParagraphs();
        }
        /// <summary>
        /// Loads an existing DOCX document and renders all paragraphs to Console.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/get-paragraphs-from-docx-document-net-csharp-vb.php
        /// </remarks>
        static void GetParagraphs()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            foreach (Paragraph par in dc.GetChildElements(true,ElementType.Paragraph))
            {
                Console.WriteLine(par.Content.ToString());
            }
            Console.ReadKey();
        }
    }
}