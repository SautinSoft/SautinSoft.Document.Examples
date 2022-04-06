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
            AddPageNumbering();            
        }

        /// <summary>
        /// Add page numbering into an existing PDF document.
        /// </summary>
        /// <remarks>
        /// https://www.sautinsoft.com/products/document/help/net/developer-guide/add-page-numbering-in-pdf-net-csharp-vb.php
        /// </remarks>
        static void AddPageNumbering()
        {
            string inpFile = @"..\..\..\shrek.pdf";
            string outFile = @"With Pages.pdf";
            DocumentCore dc = DocumentCore.Load(inpFile);

            // We place our page numbers into the footer.
            // Therefore we've to create a footer.
            HeaderFooter footer = new HeaderFooter(dc, HeaderFooterType.FooterDefault);

            // Create a new paragraph to insert a page numbering.
            // So that, our page numbering looks as: Page N of M.
            Paragraph par = new Paragraph(dc);
            par.ParagraphFormat.Alignment = HorizontalAlignment.Right;
            CharacterFormat cf = new CharacterFormat() { FontName = "Consolas", Size = 14.0 };
            par.Content.Start.Insert("Page ", cf.Clone());

            // Page numbering is a Field.
            // Create two fields: FieldType.Page and FieldType.NumPages.
            Field fPage = new Field(dc, FieldType.Page);
            fPage.CharacterFormat = cf.Clone();
            par.Content.End.Insert(fPage.Content);
            par.Content.End.Insert(" of ", cf.Clone());
            Field fPages = new Field(dc, FieldType.NumPages);
            fPages.CharacterFormat = cf.Clone();
            par.Content.End.Insert(fPages.Content);
            footer.Blocks.Add(par);

            foreach (Section s in dc.Sections)
            {
                s.HeadersFooters.Add(footer.Clone(true));
            }
            dc.Save(outFile);

            // Open the results for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(inpFile) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}