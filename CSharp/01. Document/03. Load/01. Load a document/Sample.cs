using System.IO;
using SautinSoft.Document;
using System;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            LoadFromFile();
            //LoadFromStream();
            //LoadFromBytes()
        }

        /// <summary>
        /// Loads a document into DocumentCore (dc) from a file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-document.php
        /// </remarks>
        static void LoadFromFile()
        {
            string filePath = @"..\..\..\example.docx";
            // The file format is detected automatically from the file extension: ".docx".
            // But as shown in the example below, we can specify DocxLoadOptions as 2nd parameter
            // to explicitly set that a loadable document has Docx format.
            DocumentCore dc = DocumentCore.Load(filePath);

            if (dc!=null)
                Console.WriteLine("Loaded successfully!");
			
			Console.ReadKey();			
        }

        /// <summary>
        /// Loads a document into DocumentCore (dc) from a Stream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-document.php
        /// </remarks>
        static void LoadFromStream()
        {
            // We've knowingly created an empty DocumentCore instance before "Using {}"
            // to continue work with it after stream will be closed.
            DocumentCore dc = null;
            using (FileStream fs = new FileStream(@"..\..\..\example.docx", FileMode.Open))
            {

                // Here we explicitly set that a loadable document is Docx.
                dc = DocumentCore.Load(fs, new DocxLoadOptions());
            }
            if (dc != null)
                Console.WriteLine("Loaded successfully!");
			
			Console.ReadKey();
        }

        /// <summary>
        /// Loads a document into DocumentCore (dc) from an array of bytes.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-document.php
        /// </remarks>
        static void LoadFromBytes()
        {
            // Get document bytes from a file.
            byte[] fileBytes = File.ReadAllBytes(@"..\..\..\example.pdf");

            DocumentCore dc = null;
            using (MemoryStream ms = new MemoryStream(fileBytes))
            {

                // With PdfLoadOptions we explicitly set that a loadable document is PDF.
                PdfLoadOptions pdfLO = new PdfLoadOptions()
                {

                    // 'false' - means to load vector graphics as is. Don't transform it to raster images.
                    RasterizeVectorGraphics = false,

                    // The PDF format doesn't have real tables, in fact it's a set of orthogonal graphic lines.
                    // In case of 'true' the component will detect and recreate tables from graphic lines.
                    DetectTables = false,

                    // 'Disabled' - Never load embedded fonts in PDF. Use the fonts with the same name installed at the system or similar by font metrics.
					// 'Enabled' - Always load embedded fonts in PDF.
					// 'Auto' - Load only embedded fonts missing in the system. In other case, use the system fonts.
                    PreserveEmbeddedFonts = PropertyState.Auto,

                    // Load only first 2 pages from the document.
                    PageIndex = 0,
                    PageCount = 2
                };
                dc = DocumentCore.Load(ms, pdfLO);
            }
            if (dc != null)
                Console.WriteLine("Loaded successfully!");

			Console.ReadKey();
        }
    }
}