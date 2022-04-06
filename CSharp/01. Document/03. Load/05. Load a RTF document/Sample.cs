using System;
using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
       
        static void Main(string[] args)
        {
            //LoadRtfFromStream();
            LoadRtfFromFile();
        }

        /// <summary>
        /// Loads an RTF document into DocumentCore (dc) from a file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-rtf-document-net-csharp-vb.php
        /// </remarks>
        static void LoadRtfFromFile()
        {
            string filePath = @"..\..\..\example.rtf";

            // The file format is detected automatically from the file extension: ".rtf".
            // But as shown in the example below, we can specify RtfLoadOptions as 2nd parameter
            // to explicitly set that a loadable document has RTF format.
            DocumentCore dc = DocumentCore.Load(filePath);

            if (dc != null)
                Console.WriteLine("Loaded successfully!");
			
			Console.ReadKey();			
        }

        /// <summary>
        /// Loads an RTF document into DocumentCore (dc) from a file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-rtf-document-net-csharp-vb.php
        /// </remarks>
        static void LoadRtfFromStream()
        {
            // Get document bytes.
            byte[] fileBytes = File.ReadAllBytes(@"..\..\..\example.rtf");

            DocumentCore dc = null;

            // Create a MemoryStream
            using (MemoryStream ms = new MemoryStream(fileBytes))
            {
                // Load a document from the MemoryStream.
                // Specifying RtfLoadOptions we explicitly set that a loadable document is RTF.
                dc = DocumentCore.Load(ms, new RtfLoadOptions());
            }
            if (dc != null)
                Console.WriteLine("Loaded successfully!");
			
			Console.ReadKey();			
        }
    }
}