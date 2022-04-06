using System;
using System.IO;
using System.Collections.Generic;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            MergeMultipleDocuments();
        }

		/// <summary>
        /// This sample shows how to merge multiple DOCX, RTF, PDF and Text files.
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/merge-multiple-files-net-csharp-vb.php
        /// </remarks>
        public static void MergeMultipleDocuments()
        {
            // Path to our combined document.
            string singlePDFPath = "Single.pdf";
            string workingDir = @"..\..\..\";

            List<string> supportedFiles = new List<string>();
            // Fill the collection 'supportedFiles' by *.docx, *.pdf and *.txt.
            foreach (string file in Directory.GetFiles(workingDir, "*.*"))
            {
                string ext = Path.GetExtension(file).ToLower();

                if (ext == ".docx" || ext == ".pdf" || ext == ".txt")
                    supportedFiles.Add(file);
            }


            // Create single pdf.
            DocumentCore singlePDF = new DocumentCore();

            foreach (string file in supportedFiles)
            {
                DocumentCore dc = DocumentCore.Load(file);

                Console.WriteLine("Adding: {0}...", Path.GetFileName(file));

                // Create import session.
                ImportSession session = new ImportSession(dc, singlePDF, StyleImportingMode.KeepSourceFormatting);

                // Loop through all sections in the source document.
                foreach (Section sourceSection in dc.Sections)
                {
                    // Because we are copying a section from one document to another,
                    // it is required to import the Section into the destination document.
                    // This adjusts any document-specific references to styles, bookmarks, etc.
                    //
                    // Importing a element creates a copy of the original element, but the copy
                    // is ready to be inserted into the destination document.
                    Section importedSection = singlePDF.Import<Section>(sourceSection, true, session);

                    // First section start from new page.
                    if (dc.Sections.IndexOf(sourceSection) == 0)
                        importedSection.PageSetup.SectionStart = SectionStart.NewPage;

                    // Now the new section can be appended to the destination document.
                    singlePDF.Sections.Add(importedSection);
                }
            }

            // Save single PDF to a file.
            singlePDF.Save(singlePDFPath);

            // Open the result for demonstration purposes.
           System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(singlePDFPath) { UseShellExecute = true });
        }
    }
}