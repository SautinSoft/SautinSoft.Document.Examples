using System;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System.IO;
using System.Linq;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            CalculateStatistics();
        }
		
        /// <summary>
        /// Calculates the number of words, pages and characters in a document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/counting-words-paragraphs-net-csharp-vb.php
        /// </remarks>
        static void CalculateStatistics()
        {
            // Load a DOCX file.
            string filePath = @"..\..\..\words.docx";

            DocumentCore dc = DocumentCore.Load(filePath);

            // Update and count the number of words and pages in the file.
            dc.CalculateStats();

            // Show statistics.
            Console.WriteLine("Pages: {0}", dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Pages]);
            Console.WriteLine("Paragraphs: {0}", dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Paragraphs]);
            Console.WriteLine("Words: {0}", dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Words]);
            Console.WriteLine("Characters: {0}", dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.Characters]);
            Console.WriteLine("Characters with spaces: {0}", dc.Document.Properties.BuiltIn[BuiltInDocumentProperty.CharactersWithSpaces]);
        }
    }
}