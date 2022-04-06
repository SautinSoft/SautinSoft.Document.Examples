using System;
using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {      
        static void Main(string[] args)
        {
            LoadDocxFromFile();
            //LoadDocxFromStream();
        }

        /// <summary>
        /// Loads a DOCX document into DocumentCore (dc) from a file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-docx-document-net-csharp-vb.php
        /// </remarks>
        static void LoadDocxFromFile()
        {
            string filePath = @"..\..\..\example.docx";
            // The file format is detected automatically from the file extension: ".docx".
            // But as shown in the example below, we can specify DocxLoadOptions as 2nd parameter
            // to explicitly set that a loadable document has Docx format.
            DocumentCore dc = DocumentCore.Load(filePath);

            if (dc != null)
                Console.WriteLine("Loaded successfully!");

			Console.ReadKey();
        }

        /// <summary>
        /// Loads a DOCX document into DocumentCore (dc) from a MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-docx-document-net-csharp-vb.php
        /// </remarks>
        static void LoadDocxFromStream()
        {
            // Assume that we already have a DOCX document as bytes array.
            byte[] fileBytes = File.ReadAllBytes(@"..\..\..\example.docx");

            DocumentCore dc = null;

            // Create a MemoryStream
            using (MemoryStream ms = new MemoryStream(fileBytes))
            {
                // Load a document from the MemoryStream.
                // Specifying DocxLoadOptions we explicitly set that a loadable document is Docx.
                dc = DocumentCore.Load(ms, new DocxLoadOptions());
            }
            if (dc != null)
                Console.WriteLine("Loaded successfully!");
			
			Console.ReadKey();			
        }
    }
}