using System;
using SautinSoft.Document;
using System.Text;

namespace Sample
{
    class Sample
    {
      
        static void Main(string[] args)
        {
            GetContent();
        }

		/// <summary>
        /// How to get a content from a document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/get-content-net-csharp-vb.php
        /// </remarks>
        public static void GetContent()
        {
            // Path to an input document.
            string documentPath = @"..\..\..\example.docx";

            DocumentCore dc = DocumentCore.Load(documentPath);

            StringBuilder sb = new StringBuilder();

            // Get content of each paragraph in the document.
            foreach (Paragraph par in dc.GetChildElements(true, ElementType.Paragraph))
            {
                // The property 'Content' returns the content as ContentRange.
                // Get content and append it into StringBuilder.
                sb.AppendFormat("Paragraph: {0}", par.Content.ToString());
                sb.AppendLine();
            }

            // Get content of each Run where the text color is Red.
            foreach (Run run in dc.GetChildElements(true, ElementType.Run))
            {
                if (run.CharacterFormat.FontColor == Color.Red)
                {
                    // The property 'Content' returns the content as ContentRange.
                    // Get content and append it into StringBuilder.
                    sb.AppendFormat("Red color: {0}", run.Content.ToString());
                    sb.AppendLine();
                }
            }
            Console.WriteLine(sb.ToString());
            Console.ReadKey();
        }
    }
}