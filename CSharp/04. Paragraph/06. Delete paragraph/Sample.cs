using System;
using System.IO;
using System.Linq;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            DeleteParagraphs();
        }
        /// <summary>
        /// Deletes a specific paragraphs in an existing DOCX and save it as new PDF.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/delete-paragraphs-in-docx-document-net-csharp-vb.php
        /// </remarks>
        static void DeleteParagraphs()
        {
            string filePath = @"..\..\..\example.docx";
            string fileResult = @"Result.pdf";

            DocumentCore dc = DocumentCore.Load(filePath);

            // Note, remove paragraphs only inside the first section.
            Section section = dc.Sections[0];

            // Let's remove all paragraphs containing the text "Jack".
            for (int i = 0; i < section.Blocks.Count; i++)
            {
                if (section.Blocks[i].Content.Find("Jack").Count() > 0)
                {
                    section.Blocks.RemoveAt(i);
                    i--;
                }
            }
            dc.Save(fileResult);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fileResult) { UseShellExecute = true });
        }
    }
}