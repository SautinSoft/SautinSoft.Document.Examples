using System.IO;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            LoadAndSaveAsPDFA();
        }
        
        /// <summary>
        /// Load an existing document (*.docx, *.rtf, *.pdf, *.html, *.txt, *.pdf) and save it as a PDF/A compliant version. 
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/create-and-save-document-in-pdf-a-format-net-csharp-vb.php
        /// </remarks>
        public static void LoadAndSaveAsPDFA()
        {
            // Path to a loadable document.
            string loadPath = @"..\..\..\example.docx";

            DocumentCore dc = DocumentCore.Load(loadPath);

            PdfSaveOptions options = new PdfSaveOptions()
            {
                // PdfComliance supports: PDF/A, PDF/1.5, etc.
                Compliance = PdfCompliance.PDF_A1a
            };

            string savePath = Path.ChangeExtension(loadPath, ".pdf");
            dc.Save(savePath, options);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(savePath) { UseShellExecute = true });
        }
    }
}