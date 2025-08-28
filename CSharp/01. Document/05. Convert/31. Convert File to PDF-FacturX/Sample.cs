using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            ConvertFromFile();
           // ConvertFromStream();
        }

        /// <summary>
        /// Convert RTF file to PDF/ Factur-X format (file to file).
        /// Read more information about Factur-X: https://fnfe-mpe.org/factur-x/
        /// </summary>
		/// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/convert-file-to-pdf-factur-x-in-csharp-vb.php
        /// </remarks>
        static void ConvertFromFile()
        {
            string inpFile = @"..\..\..\example.rtf";
			string xmlInfo = File.ReadAllText(@"..\..\..\info.xml");
			
            string outFile = @"..\..\..\FacturXFromRtf.pdf";

            DocumentCore dc = DocumentCore.Load(inpFile);

            PdfSaveOptions pdfSO = new PdfSaveOptions()
            {
                // Factur-X is at the same time a full readable invoice in a PDF A/3 format,
                // containing all information useful for its treatment, especially in case of discrepancy or absence of automatic matching with orders and / or receptions,
                // and a set of invoice data presented in an XML structured file conformant to EN16931 (syntax CII D16B), complete or not, allowing invoice process automation.
                FacturXXML = xmlInfo
            };

            dc.Save(outFile, pdfSO);
			
			// Important for Linux: Install MS Fonts
			// sudo apt install ttf-mscorefonts-installer -y


            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }

        /// <summary>
        /// Convert PDF file to PDF/ Factur-X format (using Stream).
        /// Read more information about Factur-X: https://fnfe-mpe.org/factur-x/
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/convert-file-to-pdf-factur-x-in-csharp-vb.php
        /// </remarks>
        static void ConvertFromStream()
        {

            // We need files only for demonstration purposes.
            // The conversion process will be done completely in memory.
            string inpFile = @"..\..\..\example.pdf";
            string xmlInfo = File.ReadAllText(@"..\..\..\info.xml");
            
            string outFile = @"..\..\..\FacturXFromPdf.pdf";
			
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
                    // 'false' - Never load embedded fonts in PDF. Use the fonts with the same name installed at the system or similar by font metrics.
                    // 'true' - Always load embedded fonts in PDF.
                    PreserveEmbeddedFonts = true
                };

                // Load a document.
                DocumentCore dc = DocumentCore.Load(msInp, pdfLO);

                // Save the document to PDF/A format.
                PdfSaveOptions pdfSO = new PdfSaveOptions()
                {
                    // Factur-X is at the same time a full readable invoice in a PDF A/3 format,
                    // containing all information useful for its treatment, especially in case of discrepancy or absence of automatic matching with orders and / or receptions,
                    // and a set of invoice data presented in an XML structured file conformant to EN16931 (syntax CII D16B), complete or not, allowing invoice process automation.
                    FacturXXML = xmlInfo
                };

                using (MemoryStream outMs = new MemoryStream())
                {
                    dc.Save(outMs, pdfSO);
					
					// Important for Linux: Install MS Fonts
					// sudo apt install ttf-mscorefonts-installer -y

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