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
            FindText();
        }
        /// <summary>
        /// Find a specific text in DOCX document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/find-text-in-docx-document-net-csharp-vb.php
        /// </remarks>
        static void FindText()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);
            string searchText = "document";
            int count = dc.Content.Find(searchText).Count();

            Console.WriteLine("The text: \"" + searchText + "\" - was found " + count.ToString() + " time(s).");
            Console.ReadKey();
        }
    }
}