using System;
using SautinSoft.Document;
using System.IO;
using System.Linq;
using System.Text;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            ImportingElement();
        }

        /// <summary>
        /// Copy a document to other document. Supported any formats (PDF, DOCX, RTF, HTML).
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/importing-element-net-csharp-vb.php
        /// </remarks>
        static void ImportingElement()
        {
            string file1 = @"..\..\..\digitalsignature.docx";
            string file2 = @"..\..\..\Parsing.docx";
            string resultFile = "Importing.docx";

            // Load files.
            DocumentCore dc = DocumentCore.Load(file1);
            DocumentCore dc1 = DocumentCore.Load(file2);

            // New Import Session to improve performance.
            var session = new ImportSession(dc1, dc);

            // Import all sections from source document.
            foreach (Section sourceSection in dc1.Sections)
            {
                Section destinationSection = dc.Import(sourceSection, true, session);
                dc.Sections.Add(destinationSection);
            }

            // Save the result.
            dc.Save(resultFile);

            // Show the result.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultFile) { UseShellExecute = true });
        }
    }
}