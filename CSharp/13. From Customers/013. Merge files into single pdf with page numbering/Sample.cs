using System;
using System.IO;
using System.Collections.Generic;
using SautinSoft.Document;
using System.Linq;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free 30-day key here:   
            // https://sautinsoft.com/start-for-free/

            MergeMultipleDocuments();
        }

        /// <summary>
        /// This sample shows how to merge DOCX, RTF, PDF, PNG or text files into a single PDF document and add page numbers inside.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-merge-multiple-files-into-single-pdf-add-page-numbering-in-csharp-vb-net.php
        /// </remarks>
        public static void MergeMultipleDocuments()
        {
            // Path to our combined document.
            string singlePDFPath = "single.pdf";
            
            List<string> supportedFiles = new List<string>();

            // Sort files by name. This way the files will be combined in alphabetical order a-, b-, c- or 1-, 2-, 3-...
            var files = Directory.GetFiles(@"..\..\..\DirToMerge\", "*").OrderByDescending(d => new FileInfo(d).Name);
            
            // Fill the collection 'supportedFiles' by *.docx, *.pdf, *.png and *.txt.
            foreach (string file in files) 
            {
                string ext = Path.GetExtension(file);

                if (ext == ".docx" || ext == ".pdf" || ext == ".txt" || ext == ".png")
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
            // We place our page numbers into the footer.
            // Therefore we've to create a footer.
            HeaderFooter footer = new HeaderFooter(singlePDF, HeaderFooterType.FooterDefault);

            // Create a new paragraph to insert a page numbering.
            // So that, our page numbering looks as: Page N of M.
            Paragraph par = new Paragraph(singlePDF);
            par.ParagraphFormat.Alignment = HorizontalAlignment.Right;
            CharacterFormat cf = new CharacterFormat() { FontName = "Consolas", Size = 18.0, FontColor = Color.Red };
            par.Content.Start.Insert("Page ", cf.Clone());

            // Page numbering is a Field.
            Field fPage = new Field(singlePDF, FieldType.Page);
            fPage.CharacterFormat = cf.Clone();
            par.Content.End.Insert(fPage.Content);
                       
            footer.Blocks.Add(par);

            foreach (Section s in singlePDF.Sections)
            {
                s.HeadersFooters.Add(footer.Clone(true));
            }

            // Save single PDF to a file.
            singlePDF.Save(singlePDFPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(singlePDFPath) { UseShellExecute = true });
        }
    }
}