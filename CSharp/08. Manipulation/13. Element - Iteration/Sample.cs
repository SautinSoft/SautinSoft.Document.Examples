using System;
using SautinSoft.Document;
using System.IO;
using System.Linq;
using System.Text;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            IterationElement();
        }

		/// <summary>
        /// Calculate sections, paragraphs, inlines, runs and fields in DOCX document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/iteration-in-element-collection-net-csharp-vb.php
        /// </remarks>
        static void IterationElement()
        {
            DocumentCore dc = DocumentCore.Load(@"..\..\..\Parsing.docx", LoadOptions.DocxDefault);
            int numberOfSections = dc.Sections.Count;
            int numberOfParagraphs = dc.GetChildElements(true, ElementType.Paragraph).Count();
            int numberOfRunsAndFields = dc.GetChildElements(true, ElementType.Run, ElementType.Field).Count();
            int numberOfInlines = dc.GetChildElements(true).OfType<Inline>().Count();
            int elements = dc.Sections[0].GetChildElements(true).Count();
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("File has:");
            sb.AppendLine(numberOfSections + " section");
            sb.AppendLine(numberOfParagraphs + " paragraphs");
            sb.AppendLine(numberOfRunsAndFields + " runs and fields");
            sb.AppendLine(numberOfInlines + " inlines");
            sb.AppendLine("First section contains " + elements + " elements");
            Console.WriteLine(sb.ToString());
            Console.ReadKey();
        }
    }
}