using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {

        static void Main(string[] args)
        {
            SaveToPdfFile();
            SaveToPdfStream();
        }

        /// <summary>
        /// Creates a new document and saves it as PDF file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-pdf-net-csharp-vb.php
        /// </remarks>
        static void SaveToPdfFile()
        {
            // Assume we already have a document 'dc'.
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert("Hey Guys and Girls!\nFrom file.", new CharacterFormat() { FontColor = Color.Green, Size = 20});

            string filePath = @"Result-file.pdf";

            dc.Save(filePath, new PdfSaveOptions()
            {
                Compliance = PdfCompliance.PDF_A1a,
                PreserveFormFields = true
            });

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }

        /// <summary>
        /// Creates a new document and saves it as PDF/A using MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-pdf-net-csharp-vb.php
        /// </remarks>
        static void SaveToPdfStream()
        {
            // There variables are necessary only for demonstration purposes.
            byte[] fileData = null;
            string filePath = @"Result-stream.pdf";

            // Assume we already have a document 'dc'.
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert("Hey Guys and Girls!\nFrom MemoryStream.", new CharacterFormat() { FontColor = Color.Orange, Size = 20 });

            // Let's save our document to a MemoryStream.
            using (MemoryStream ms = new MemoryStream())
            {
                dc.Save(ms, new PdfSaveOptions()
                {
                    PageIndex = 0,
                    PageCount = 1,
                    Compliance = PdfCompliance.PDF_A1a
                });
                fileData = ms.ToArray();
            }
            File.WriteAllBytes(filePath, fileData);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });

        }
    }
}