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
            MergeDocumentsInMem();
        }

        /// <summary>
        /// This sample shows how to merge multiple files DOCX, PDF into single document in memory.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/merge-multiple-documents-docx-pdf-in-memory-net-csharp-vb.php
        /// </remarks>
        public static void MergeDocumentsInMem()
        {
            // We'll use these files only to get data and show the result.
            // The whole merging process will be done in memory.
            string[] documents = new string[] { @"..\..\..\one.docx", @"..\..\..\two.pdf" };
            string singleDocumentPath = "Result.pdf";

            // Read the files and retrieve the file data into this Dictionary.
            // Thus we'll have input data completely in memory.
            Dictionary<string, byte[]> documentsData = new Dictionary<string, byte[]>();
            foreach (string file in documents)
            {
                documentsData.Add(file, File.ReadAllBytes(file));
            }

            // Merge documents in memory (using MemoryStream)
            // 1. Create a single document.
            DocumentCore dcSingle = new DocumentCore();

            foreach (KeyValuePair<string, byte[]> document in documentsData)
            {
                // Create new MemoryStream based on document byte array.
                using (MemoryStream msDoc = new MemoryStream(document.Value))
                {
                    LoadOptions lo = null;
                    if (Path.GetExtension(document.Key).ToLower() == ".docx")
                        lo = new DocxLoadOptions();
                    else if (Path.GetExtension(document.Key).ToLower() == ".pdf")
                        lo = new PdfLoadOptions() 
						{
							// 'Disabled' - Never load embedded fonts in PDF. Use the fonts with the same name installed at the system or similar by font metrics.
							// 'Enabled' - Always load embedded fonts in PDF.
							// 'Auto' - Load only embedded fonts missing in the system. In other case, use the system fonts.
							PreserveEmbeddedFonts = PropertyState.Auto 
						};

                    DocumentCore dc = DocumentCore.Load(msDoc, lo);

                    Console.WriteLine("Adding: {0}...", Path.GetFileName(document.Key));

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
            }

            // Save the resulting document as PDF into MemoryStream.
            using (MemoryStream msPdf = new MemoryStream())
            {
                dcSingle.Save(msPdf, new PdfSaveOptions());

                // Let's also save our document to PDF file for showing the result.
                File.WriteAllBytes(singleDocumentPath, msPdf.ToArray());

                // Open the result for demonstration purposes.
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(singleDocumentPath) { UseShellExecute = true });
            }
        }
    }
}