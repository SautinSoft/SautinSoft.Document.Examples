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
            GetText();
        }
        /// <summary>
        /// Get all Text (Run objects) from DOCX document and show it on Console.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/get-text-from-docx-document-net-csharp-vb.php
        /// </remarks>
        static void GetText()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);

            // Get all Run elements from document.
            foreach (Run run in dc.GetChildElements(true,ElementType.Run))
            {
                Console.WriteLine(run.Text);
            }

            Console.ReadKey();
        }
    }
}