using SautinSoft.Document;

namespace Example
{
    class Program
    {        
        static void Main(string[] args)
        {
            ScannedPdfToWord();
        }

        /// <summary>
        /// The method converts a PDF document with scanned images to Word. But it works only if the PDF document contains a hidden text atop of the images.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-scanned-pdf-to-word-in-csharp-vb-net.php
        /// </remarks>
        static void ScannedPdfToWord()
        {
            // Actually there are a lot of PDF documents which looks like created using a scanner, 
            // but they also contain a hidden text atop of the contents. 
            // This hidden text duplicates the content of the scanned images. 
            // This is made specially to have the ability to perform the 'find' operation.

            // Our steps:
            // 1. Load the PDF with the these settings: 
            // - show hidden text;
            // - skip all images during the loading. 
            // 2. Change the font color to the 'Black' for the all text.
            // 3. Save the document as DOCX.
            string inpFile = @"..\..\..\Scanned.pdf";
            string outFile = @"Result.docx";

            PdfLoadOptions pdfLO = new PdfLoadOptions()
            {
				// 'Disabled' - Never load embedded fonts in PDF. Use the fonts with the same name installed at the system or similar by font metrics.
				// 'Enabled' - Always load embedded fonts in PDF.
				// 'Auto' - Load only embedded fonts missing in the system. In other case, use the system fonts.
                PreserveEmbeddedFonts = PropertyState.Enabled,
                PreserveImages = false,
                ShowInvisibleText = true,                
            };

            DocumentCore dc = DocumentCore.Load(inpFile, pdfLO);

            dc.DefaultCharacterFormat.FontColor = Color.Black;
            foreach (Element element in dc.GetChildElements(true, ElementType.Paragraph))
            {
                foreach (Inline inline in (element as Paragraph).Inlines)
                {
                    if (inline is Run)
                        (inline as Run).CharacterFormat.FontColor = Color.Black;
                }
                (element as Paragraph).CharacterFormatForParagraphMark.FontColor = Color.Black;
            }
            dc.Save(outFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}