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
            PageNumbering();            
        }

		/// <summary>
        /// Creates a new document with page numbering: Page N of M.
        /// </summary>
        /// <remarks>
        /// https://sautinsoft.com/products/document/help/net/developer-guide/page-numbering.php
        /// </remarks>
        public static void PageNumbering()
        {
            string documentPath = @"PageNumbering.docx";

            // Let's create a new document with multiple pages.
            DocumentCore dc = new DocumentCore();

            string[] pagesText = new string[] { "One", "Two", "Three", "Four", "Five" };
            Random r = new Random();

            // Create a new section.
            Section section = new Section(dc);
            dc.Sections.Add(section);

            // We place our page numbers into the footer.
            // Therefore we've to create a footer.
            HeaderFooter footer = new HeaderFooter(dc, HeaderFooterType.FooterDefault);

            // Create a new paragraph to insert a page numbering.
            // So that, our page numbering looks as: Page N of M.
            Paragraph par = new Paragraph(dc);
            par.ParagraphFormat.Alignment = HorizontalAlignment.Left;
            CharacterFormat cf = new CharacterFormat() { FontName = "Arial", Size = 12.0 };
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

            // Add the paragraph with Fields into the footer.
            footer.Blocks.Add(par);

            // Add the footer into the section.
            section.HeadersFooters.Add(footer);

            // Add some paragraphs with page breaks in the document,
            // to make several pages.
            foreach (string text in pagesText)
            {
                Paragraph p = new Paragraph(dc);
                p.ParagraphFormat.Alignment = HorizontalAlignment.Center;
                section.Blocks.Add(p);

                string color = String.Format("#{0:X2}{1:X2}{2:X2}", r.Next(0,255),r.Next(0,255),r.Next(0,255));
                p.Content.Start.Insert(text, new CharacterFormat() { FontName = "Arial", Size = 72.0, FontColor = new Color(color) });
                
                if (text!=pagesText.Last())
                    p.Content.End.Insert(new SpecialCharacter(dc, SpecialCharacterType.PageBreak).Content);

            }

            // Save our document into DOCX format.
            dc.Save(documentPath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}