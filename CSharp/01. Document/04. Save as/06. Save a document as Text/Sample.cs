using System;
using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
       
        static void Main(string[] args)
        {
            SaveToTextFile();
            SaveToTextStream();
        }

        /// <summary>
        /// Creates a new document and saves it as Text file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-text-net-csharp-vb.php
        /// </remarks>
        static void SaveToTextFile()
        {
            // Assume we already have a document 'dc'.
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert("Hey Guys and Girls!");

            string filePath = @"Result-file.txt";

            dc.Save(filePath, new TxtSaveOptions()
            {
                Encoding = System.Text.Encoding.UTF8,
                ParagraphBreak = Environment.NewLine
            });

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }

        /// <summary>
        /// Creates a new document and saves it as Text using MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-text-net-csharp-vb.php
        /// </remarks>
        static void SaveToTextStream()
        {
            // There variables are necessary only for demonstration purposes.
            byte[] fileData = null;
            string filePath = @"Result-stream.txt";

            // Assume we already have a document 'dc'.
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert("Hey Guys and Girls!");

            // Let's save our document to a MemoryStream.
            using (MemoryStream ms = new MemoryStream())
            {
                dc.Save(ms, new TxtSaveOptions());
                fileData = ms.ToArray();
            }
            File.WriteAllBytes(filePath, fileData);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}