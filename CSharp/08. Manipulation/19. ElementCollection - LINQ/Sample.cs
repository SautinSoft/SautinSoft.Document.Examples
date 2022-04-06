using System;
using System.Linq;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            ShowLists();
        }
        /// <summary>
        /// Find all paragraphs in a document marked as list (ordered or unordered).
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/elementcollection-linq.php
        /// </remarks>
        static void ShowLists()
        {
            string filePath = @"..\..\..\example.docx";
            DocumentCore dc = DocumentCore.Load(filePath);

            // Select all paragraphs marked as list using LINQ.
            var selectedPars = from p in dc.GetChildElements(true, ElementType.Paragraph)
                               where (p as Paragraph).ListFormat.IsList
                               select p;

            foreach (Paragraph p in selectedPars)
                Console.WriteLine(p.Content.ToString().TrimEnd());

            Console.ReadKey();
        }
    }
}