using System;
using System.Collections.Generic;
using System.Linq;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            DeletePageNumbering();
        }
        /// <summary>
        /// How to delete the page numbering in an existing DOCX document. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/remove-page-numbering-in-docx-document-net-csharp-vb.php
        /// </remarks>
        public static void DeletePageNumbering()
        {
            string inpFile = @"..\..\..\PageNumbering.docx";
            // Save the output document into PDF format,
            // you may change it to any from: DOCX, HTML, PDF, RTF.
            string outFile = @"PageNumbering - deleted.pdf";

            // Load a document from DOCX file (PageNumbering.docx).
            // We've created PageNumbering.docx in the previous example.
            DocumentCore dc = DocumentCore.Load(inpFile);
            Section s = dc.Sections[0];

            foreach (HeaderFooter hf in s.HeadersFooters)
            {
                foreach (Field field in hf.GetChildElements(true, ElementType.Field).Reverse())
                {
                    // Page numbering is a Field,
                    // so we have to find the fields with the type Page or NumPages and remove.
                    if (field.FieldType == FieldType.Page || field.FieldType == FieldType.NumPages)
                    {
                        // Also assume that our page numbering located in a paragraph,
                        // so let's remove the paragraph's content too.
                        if (field.Parent is Paragraph)
                            (field.Parent as Paragraph).Inlines.Clear();

                        // If we'll delete only the fields (field.Content.Delete()), our paragraph
                        // may contain text "Page of ".
                        // Based on this, we remove the whole paragraph.
                    }
                }
            }
            dc.Save(outFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(inpFile) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}