using System;
using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
       
        static void Main(string[] args)
        {
            LoadPDFFromFile();
            //LoadPDFFromStream();
        }

        /// <summary>
        /// Loads a PDF document into DocumentCore (dc) from a file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-pdf-document-net-csharp-vb.php
        /// </remarks>
        static void LoadPDFFromFile()
        {
            string filePath = @"..\..\..\example.pdf";

            // The file format is detected automatically from the file extension: ".pdf".
            // But as shown in the example below, we can specify PdfLoadOptions as 2nd parameter
            // to explicitly set that a loadable document has PDF format.
            DocumentCore dc = DocumentCore.Load(filePath);

            if (dc != null)
                Console.WriteLine("Loaded successfully!");
			
			Console.ReadKey();			
        }

        /// <summary>
        /// Loads a PDF document into DocumentCore (dc) from a MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/load-pdf-document-net-csharp-vb.php
        /// </remarks>
        static void LoadPDFFromStream()
        {
            // Assume that we already have a PDF document as bytes array.
            byte[] fileBytes = File.ReadAllBytes(@"..\..\..\example.pdf");

            DocumentCore dc = null;

            // Create a MemoryStream
            using (MemoryStream pdfStream = new MemoryStream(fileBytes))
            {
                // Specifying PdfLoadOptions we explicitly set that a loadable document is PDF.
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

                    // Load only the 1st page from the document.
                    PageIndex = 0,
                    PageCount = 1
                };

                // Load a PDF document from the MemoryStream.
                dc = DocumentCore.Load(pdfStream, new PdfLoadOptions());
            }
            if (dc != null)
                Console.WriteLine("Loaded successfully!");
			
			Console.ReadKey();			
        }
    }
}