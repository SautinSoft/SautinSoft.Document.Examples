using System;
using System.Linq;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            TOC();
        }

        /// <summary>
        /// Create a document with table of content.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/table-of-contents-toc.php
        /// </remarks>
        public static void TOC()
        {
            string resultFile = "Table-Of-Contents.docx";

            // Let's create a simple document.
            DocumentCore dc = new DocumentCore();

            // Create and add Heading 1 style for TOC.
            ParagraphStyle Heading1Style = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Heading1, dc);
            dc.Styles.Add(Heading1Style);

            // Create and add Heading 2 style for TOC.
            ParagraphStyle Heading2Style = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Heading2, dc);
            dc.Styles.Add(Heading2Style);

            // Add new section.
            Section section = new Section(dc);
            dc.Sections.Add(section);

            // Add TOC title in the DOCX document.
            section.Blocks.Add(new Paragraph(dc, "Table of Contents"));

            // Create and add new TOC.
            section.Blocks.Add(new TableOfEntries(dc, FieldType.TOC));

            section.Blocks.Add(new Paragraph(dc, "The end."));

            // Let's add a page break into our paragraph.
            section.Blocks.Add(
                    new Paragraph(dc,
                    new SpecialCharacter(dc, SpecialCharacterType.PageBreak)));

            // Add document content.
            // Add Chapter 1
            section.Blocks.Add(
                new Paragraph(dc, "Chapter 1")
                {
                    ParagraphFormat =
                    {
                Style = Heading1Style
                    }
                });

            // Add SubChapter 1-1
            section.Blocks.Add(
                        new Paragraph(dc, String.Format("Subchapter 1-1"))
                        {
                            ParagraphFormat =
                            {
                    Style = Heading2Style
                            }
                        });
            // Add the content of Chapter 1 / Subchapter 1-1
            section.Blocks.Add(
                        new Paragraph(dc,
                            "«Document .Net» will help you in development of applications which operates with DOCX, RTF, PDF and Text documents. After adding of the reference to (SautinSoft.Document.dll) - it's 100% C# managed assembly you will be able to create a new document, parse an existing, modify anything what you want."));

            // Let's add an another page break into our paragraph.
            section.Blocks.Add(
                   new Paragraph(dc,
                   new SpecialCharacter(dc, SpecialCharacterType.PageBreak)));

            // Add document content.
            // Add Chapter 2
            section.Blocks.Add(
                     new Paragraph(dc, "Chapter 2")
                     {
                         ParagraphFormat =
                         {
                Style = Heading1Style
                         }
                     });

            // Add SubChapter 2-1
            section.Blocks.Add(
                new Paragraph(dc, String.Format("Subchapter 2-1"))
                {
                    ParagraphFormat =
                    {
                    Style = Heading2Style
                    }
                });

            // Add the content of Chapter 2 / Subchapter 2-1
            section.Blocks.Add(
                        new Paragraph(dc,
                            "Requires only .Net 4.0 or above. Our product is compatible with all .Net languages and supports all Operating Systems where .Net Framework can be used. Note that «Document .Net» is entirely written in managed C#, which makes it absolutely standalone and an independent library. Of course, No dependency on Microsoft Word."));

            // Update TOC (TOC can be updated only after all document content is added).
            var toc = (TableOfEntries)dc.GetChildElements(true, ElementType.TableOfEntries).FirstOrDefault();
            toc.Update();

            // Update TOC's page numbers.
            // Page numbers are automatically updated in that case.
            dc.GetPaginator(new PaginatorOptions() { UpdateFields = true });

            // Save DOCX to a file
            dc.Save(resultFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultFile) { UseShellExecute = true });
        }
    }
}