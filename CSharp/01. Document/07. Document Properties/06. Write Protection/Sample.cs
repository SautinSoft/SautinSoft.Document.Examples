using System;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System.IO;
using System.Linq;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            WriteProtection();
        }

        /// <summary>
        /// Create a write protected DOCX document.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/write-protection-options-net-csharp-vb.php
        /// </remarks>
        public static void WriteProtection()
        {
            string filePath = @"ProtectedDocument.docx";

            DocumentCore dc = new DocumentCore();

            // Insert paragraphs into the document.
            dc.Sections.Add(
           new Section(dc,
               new Paragraph(dc, "This document has been opened in read only mode."),
               new Paragraph(dc, "To keep your changes, you 'll need to save the document with a new name or in a different location."),
               new Paragraph(dc, "To make changes to the current document, restart with the password '12345'.")));

            // Sets the write protection password "12345".
            DocumentWriteProtection protection = dc.WriteProtection;
            protection.SetPassword("12345");

            // Save a document as the DOCX file with write protection options.
            dc.Save(filePath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}
