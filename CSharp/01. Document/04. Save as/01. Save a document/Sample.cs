using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        
        static void Main(string[] args)
        {
            SaveToFile();
            SaveToStream();
        }
		
        /// <summary>
        /// Creates a new document and saves it as DOCX file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document.php
        /// </remarks>
        static void SaveToFile()
        {
            // Assume we already have a document 'dc'.
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert("Hey Guys and Girls!");

            string filePath = @"Result.docx";

            // The file format will be detected automatically from the file extension: ".docx".
            dc.Save(filePath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }

        /// <summary>
        /// Creates a new document and saves it as PDF using MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document.php
        /// </remarks>
        static void SaveToStream()
        {
            // There variables are necessary only for demonstration purposes.
            byte[] fileData = null;
            string filePath = @"Result.pdf";

            // Assume we already have a document 'dc'.
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert("Hey Guys and Girls!");

            // Let's save our document to a MemoryStream.
            using (MemoryStream ms = new MemoryStream())
            {
                // 2nd parameter: we've explicitly set to save our document in PDF format.
                dc.Save(ms, new PdfSaveOptions());

                fileData = ms.ToArray();
            }

            File.WriteAllBytes(filePath, fileData);
            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}