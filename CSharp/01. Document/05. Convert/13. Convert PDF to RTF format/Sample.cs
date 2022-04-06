using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            ConvertFromFile();
            ConvertFromStream();
        }

        /// <summary>
        /// Convert PDF to RTF (file to file).
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-pdf-to-rtf-in-csharp-vb.php
        /// </remarks>
        static void ConvertFromFile()
        {
            string inpFile = @"..\..\..\example.pdf";
            string outFile = @"Result.rtf";

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
                PreserveEmbeddedFonts = PropertyState.Auto
            };

            DocumentCore dc = DocumentCore.Load(inpFile, pdfLO);
            dc.Save(outFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }

        /// <summary>
        /// Convert PDF to RTF (using Stream).
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-pdf-to-rtf-in-csharp-vb.php
        /// </remarks>
        static void ConvertFromStream()
        {

            // We need files only for demonstration purposes.
            // The conversion process will be done completely in memory.
            string inpFile = @"..\..\..\example.pdf";
            string outFile = @"ResultStream.rtf";
            byte[] inpData = File.ReadAllBytes(inpFile);
            byte[] outData = null;

            using (MemoryStream msInp = new MemoryStream(inpData))
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
                    PreserveEmbeddedFonts = PropertyState.Auto
                };

                // Load a document.
                DocumentCore dc = DocumentCore.Load(msInp, pdfLO);

                // Save the document to RTF format.
                using (MemoryStream outMs = new MemoryStream())
                {
                    dc.Save(outMs, new RtfSaveOptions() );
                    outData = outMs.ToArray();                    
                }
                // Show the result for demonstration purposes.
                if (outData != null)
                {
                    File.WriteAllBytes(outFile, outData);
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
                }
            }
        }
    }
}