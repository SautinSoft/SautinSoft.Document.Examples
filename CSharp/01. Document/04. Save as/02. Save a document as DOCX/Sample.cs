using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
       
        static void Main(string[] args)
        {
            SaveToDocxFile();
            SaveToDocxStream();
        }

        /// <summary>
        /// Creates a new document and saves it as DOCX file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-docx-net-csharp-vb.php
        /// </remarks>
        static void SaveToDocxFile()
        {
            // Assume we already have a document 'dc'.
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert("Hey from File!");

            string filePath = @"Result-file.docx";

            dc.Save(filePath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });

        }

        /// <summary>
        /// Creates a new document and saves it as DOCX using MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-docx-net-csharp-vb.php
        /// </remarks>
        static void SaveToDocxStream()
        {
            // There variables are necessary only for demonstration purposes.
            byte[] fileData = null;
            string filePath = @"Result-stream.docx";

            // Assume we already have a document 'dc'.
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert("Hey from MemoryStream!");

            // Let's save our document to a MemoryStream.
            using (MemoryStream ms = new MemoryStream())
            {
                dc.Save(ms, new DocxSaveOptions());
                fileData = ms.ToArray();
            }
            File.WriteAllBytes(filePath, fileData);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}