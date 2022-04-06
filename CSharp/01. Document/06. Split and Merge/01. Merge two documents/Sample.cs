using System;
using System.IO;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            MergeTwoDocuments();
        }
        
		/// <summary>
        /// How to merge two documents: DOCX and PDF.
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/split-and-merge-content-net-csharp-vb.php
        /// </remarks>
        public static void MergeTwoDocuments()
        {
            // Path to our combined document.
            string singleFilePath = "Single.docx";

            string[] supportedFiles = new string[] { @"..\..\..\example.docx", @"..\..\..\example.pdf" };

            // Create single document.
            DocumentCore dcSingle = new DocumentCore();

            foreach (string file in supportedFiles)
            {
                DocumentCore dc = DocumentCore.Load(file);

                Console.WriteLine("Adding: {0}...", Path.GetFileName(file));

                // Create import session.
                ImportSession session = new ImportSession(dc, dcSingle, StyleImportingMode.KeepSourceFormatting);

                // Loop through all sections in the source document.
                foreach (Section sourceSection in dc.Sections)
                {
                    // Because we are copying a section from one document to another,
                    // it is required to import the Section into the destination document.
                    // This adjusts any document-specific references to styles, bookmarks, etc.
                    //
                    // Importing a element creates a copy of the original element, but the copy
                    // is ready to be inserted into the destination document.
                    Section importedSection = dcSingle.Import<Section>(sourceSection, true, session);

                    // First section start from new page.
                    if (dc.Sections.IndexOf(sourceSection) == 0)
                        importedSection.PageSetup.SectionStart = SectionStart.NewPage;

                    // Now the new section can be appended to the destination document.
                    dcSingle.Sections.Add(importedSection);
                }
            }

            // Save single document to a file.
            dcSingle.Save(singleFilePath);

            // Open the result for demonstration purposes.
           System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(singleFilePath) { UseShellExecute = true });
        }
    }
}