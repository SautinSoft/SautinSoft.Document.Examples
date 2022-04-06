using System;
using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {      
        static void Main(string[] args)
        {
            LoadDocFromFile();
            //LoadDocFromStream();
        }

        /// <summary>
        /// Loads a DOC (Word 97-2003) document into DocumentCore (dc) from a file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-doc-word-97-2003-document-net-csharp-vb.php
        /// </remarks>
        static void LoadDocFromFile()
        {
            string filePath = @"..\..\..\example.doc";
            // The file format is detected automatically from the file extension: ".doc".
            // But as shown in the example below, we can specify DocLoadOptions as 2nd parameter
            // to explicitly set that a loadable document has DOC format.
            DocumentCore dc = DocumentCore.Load(filePath);

            if (dc != null)
                Console.WriteLine("Loaded successfully!");

			Console.ReadKey();
        }

        /// <summary>
        /// Loads a DOC (Word 97-2003) document into DocumentCore (dc) from a MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-doc-word-97-2003-document-net-csharp-vb.php
        /// </remarks>
        static void LoadDocFromStream()
        {
            // Assume that we already have a DOC (Word 97-2003) document as bytes array.
            byte[] fileBytes = File.ReadAllBytes(@"..\..\..\example.doc");

            DocumentCore dc = null;

            // Create a MemoryStream
            using (MemoryStream ms = new MemoryStream(fileBytes))
            {
                // Load a document from the MemoryStream.
                // Specifying DocLoadOptions we explicitly set that a loadable document is DOC.
                dc = DocumentCore.Load(ms, new DocLoadOptions());
            }
            if (dc != null)
                Console.WriteLine("Loaded successfully!");
			
			Console.ReadKey();			
        }
    }
}