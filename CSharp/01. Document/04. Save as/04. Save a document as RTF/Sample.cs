using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {

        static void Main(string[] args)
        {
            SaveToRtfFile();
            SaveToRtfStream();
        }

        /// <summary>
        /// Creates a new document and saves it as RTF file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-rtf-net-csharp-vb.php
        /// </remarks>
        static void SaveToRtfFile()
        {
            // Assume we already have a document 'dc'.
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert("Hey Guys and Girls!\nFrom file.", new CharacterFormat() { FontColor = Color.Green, Size = 20 });

            string filePath = @"Result-file.rtf";

            dc.Save(filePath, new RtfSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }

        /// <summary>
        /// Creates a new document and saves it as RTF using MemoryStream.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-rtf-net-csharp-vb.php
        /// </remarks>
        static void SaveToRtfStream()
        {
            // There variables are necessary only for demonstration purposes.
            byte[] fileData = null;
            string filePath = @"Result-stream.rtf";

            // Assume we already have a document 'dc'.
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert("Hey Guys and Girls!\nFrom MemoryStream.", new CharacterFormat() { FontColor = Color.Orange, Size = 20 });

            // Let's save our document to a MemoryStream.
            using (MemoryStream ms = new MemoryStream())
            {
                dc.Save(ms, new RtfSaveOptions());
                fileData = ms.ToArray();
            }
            File.WriteAllBytes(filePath, fileData);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}