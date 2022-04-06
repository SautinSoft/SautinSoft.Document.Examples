using System;
using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        
        static void Main(string[] args)
        {
            LoadHtmlFromFile();
            //LoadHtmlFromStream();
        }

        /// <summary>
        /// Loads an HTML document into DocumentCore (dc) from a file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-html-document-net-csharp-vb.php
        /// </remarks>
        static void LoadHtmlFromFile()
        {
            string filePath = @"..\..\..\example.html";
            // The file format is detected automatically from the file extension: ".html".
            // But as shown in the example below, we can specify HtmlLoadOptions as 2nd parameter
            // to explicitly set that a loadable document has HTML format.
            DocumentCore dc = DocumentCore.Load(filePath);
            if (dc != null)
                Console.WriteLine("Loaded successfully!");

			Console.ReadKey();			
        }

        /// <summary>
        /// Loads an HTML document into DocumentCore (dc) from a MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-html-document-net-csharp-vb.php
        /// </remarks>
        static void LoadHtmlFromStream()
        {
            // Get document bytes.
            byte[] fileBytes = File.ReadAllBytes(@"..\..\..\example.html");

            DocumentCore dc = null;

            // Create a MemoryStream
            using (MemoryStream ms = new MemoryStream(fileBytes))
            {
                // Load a document from the MemoryStream.
                // Specifying HtmlLoadOptions we explicitly set that a loadable document is HTML.
                dc = DocumentCore.Load(ms, new HtmlLoadOptions());
            }
            if (dc != null)
                Console.WriteLine("Loaded successfully!");
			
			Console.ReadKey();			
        }
    }
}