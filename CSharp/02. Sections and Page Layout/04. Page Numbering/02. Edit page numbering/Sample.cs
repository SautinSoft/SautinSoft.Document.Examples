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
            EditPageNumbering();
        }

        /// <summary>
        /// How to edit Page Numbering in an existing DOCX document. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/edit-page-numbering-in-docx-document-net-csharp-vb.php
        /// </remarks>
        public static void EditPageNumbering()
        {
            // Modify the Page Numbering from: "Page N of M" to "Página N".
            string inpFile = @"..\..\..\PageNumbering.docx";
            string outFile = @"PageNumbering-modified.pdf";

            // Load a DOCX document with Page Numbering (PageNumbering.docx).
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
                        // Save the paragraph containing the page numbering and character formatting.
                        Paragraph parWithNumbering = null;
                        CharacterFormat cf = field.CharacterFormat.Clone();

                        // Also assume that our page numbering located in a paragraph,
                        // so let's remove the paragraph's content too.
                        if (field.Parent is Paragraph)
                        {
                            parWithNumbering = (field.Parent as Paragraph);
                            // If we'll delete only the fields (field.Content.Delete()), our paragraph
                            // may contain text "Page of ".
                            // Based on this, we remove all items in this paragraph.
                            // Fields {Page} and {NumPages} will be also deleted here, because they are Inlines too.
                            parWithNumbering.Inlines.Clear();
                        }

                        if (parWithNumbering != null)
                        {
                            // Insert new Page Numbering: "Página N":                        
                            // Insert "Página ".
                            parWithNumbering.Inlines.Add(new Run(dc, "Página ", cf.Clone()));
                            // Insert field {Page}.
                            Field fPage = new Field(dc, FieldType.Page);
                            fPage.CharacterFormat = cf.Clone();
                            parWithNumbering.Inlines.Add(fPage);
                        }
                    }
                }
            }
            dc.Save(outFile);

            // Open the results for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(inpFile) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}