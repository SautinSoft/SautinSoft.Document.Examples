using System;
using System.IO;
using System.Collections.Generic;
using SautinSoft.Document;
using SautinSoft;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            ConvertToSingleXls();
        }

        /// <summary>
        /// How to convert all files to a single XLS file.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-convert-pdf-docx-rtf-to-single-xls-workbook-net-csharp-vb.php
        /// </remarks>
        public static void ConvertToSingleXls()
        {
            // In this example we'll use not only Document .Net component, but also
            // another SautinSoft 'component - PDF Focus .Net (to perform conversion from PDF to single xls workbook).
			// First of all, please perform "Rebuild Solution" to restore PDF Focus .Net package from NuGet.

            // Our steps:
            // 1. Convert all RTF, DOCX, PDF files into a single PDF document. (by Document .Net).
            // 2. Convert the single PDF into a single XLS workbook. (by PDF Focus .Net).

            byte[] singlePdfBytes = null;

            // This file we need only to show intermediate result.
            string singlePdfFile = "Single.pdf";
            string workingDir = @"..\..\..\";
            string singleXlsFile = "Single.xls";

            List<string> supportedFiles = new List<string>();

            foreach (string file in Directory.GetFiles(workingDir, "*.*"))
            {
                string ext = Path.GetExtension(file).ToLower();

                if (ext == ".pdf" || ext == ".docx" || ext == ".rtf")
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

            // Save our single document into PDF format in memory.
            // Let's save our document to a MemoryStream.
            using (MemoryStream Pdf = new MemoryStream())
            {
                singlePDF.Save(Pdf, new PdfSaveOptions()
                {
                    Compliance = PdfCompliance.PDF_A1a
                });
                singlePdfBytes = Pdf.ToArray();
            }

            // Open the result for demonstration purposes.
            File.WriteAllBytes(singlePdfFile, singlePdfBytes);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(singlePdfFile) { UseShellExecute = true });

            SautinSoft.PdfFocus f = new PdfFocus();
          
            f.OpenPdf(singlePdfBytes);

            if (f.PageCount > 0)
                f.ToExcel(singleXlsFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(singleXlsFile) { UseShellExecute = true });
        }
    }
}