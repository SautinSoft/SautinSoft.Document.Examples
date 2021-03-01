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
            InsertText();
        }
        /// <summary>
        /// Insert a text into an existing PDF document in a specific position.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-text-to-pdf-document-net-csharp-vb.php
        /// </remarks>
        static void InsertText()
        {
            string filePath = @"..\..\example.pdf";
            string fileResult = @"Result.pdf";
            DocumentCore dc = DocumentCore.Load(filePath);

            // Find a position to insert text. Before this text: "> in this position".
            ContentRange cr =  dc.Content.Find("> in this position").FirstOrDefault();

            // Insert new text.
            if (cr != null)
                cr.Start.Insert("New text!");
            dc.Save(fileResult);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(fileResult) { UseShellExecute = true });

        }
    }
}