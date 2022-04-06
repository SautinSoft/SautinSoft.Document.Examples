using System;
using System.Linq;
using SautinSoft.Document;

namespace Example
{
    class Sample
    {
        static void Main(string[] args)
        {
            ExtendedTOC();
        }

        /// <summary>
        /// Create extended table of contents in word document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-extended-table-of-contents-in-word-document-net-csharp-vb.php
        /// </remarks> 
        public static void ExtendedTOC()
        {
            string resultFile = "Extended-Table-Of-Contents.docx";
            // First of all, create an instance of DocumentCore.
            DocumentCore dc = new DocumentCore();

            // Create and add Heading1 style. For "Chapter 1" and "Chapter 2".
            ParagraphStyle Heading1Style = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Heading1, dc);
            Heading1Style.ParagraphFormat.LineSpacing = 3;
            Heading1Style.CharacterFormat.Size = 18;
            // #358CCB - blue
            Heading1Style.CharacterFormat.FontColor = new Color("#358CCB");
            dc.Styles.Add(Heading1Style);

            // Create and add Heading2 style. For "SupChapter 1-1" and "SubChapter 2-1".
            ParagraphStyle Heading2Style = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Heading2, dc);
            Heading2Style.ParagraphFormat.LineSpacing = 2;
            Heading2Style.CharacterFormat.Size = 14;
            // #FF9900 - orange
            Heading2Style.CharacterFormat.FontColor = new Color("#FF9900");
            dc.Styles.Add(Heading2Style);

            // Create and add TOC style.
            ParagraphStyle TOCStyle = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Subtitle, dc);
            TOCStyle.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText;
            TOCStyle.ParagraphFormat.Alignment = HorizontalAlignment.Center;
            TOCStyle.CharacterFormat.Bold = true;
            // #358CCB - blue
            TOCStyle.CharacterFormat.FontColor = new Color("#358CCB");
            dc.Styles.Add(TOCStyle);

            // Add new section.
            Section section = new Section(dc);
            dc.Sections.Add(section);

            // Add TOC Header.
            section.Blocks.Add(
                new Paragraph(dc, "Table of Contents")
                { ParagraphFormat = { Style = TOCStyle } });

            // Create and add TOC (Table of Contents).
            section.Blocks.Add(new TableOfEntries(dc, FieldType.TOC));

            // Add TOC Ending.
            section.Blocks.Add(
                 new Paragraph(dc, "The End")
                 { ParagraphFormat = { Alignment = HorizontalAlignment.Center, BackgroundColor = Color.Gray } });

            // Add the document content (Chapter 1).
            // Add Chapter 1.
            section.Blocks.Add(
                new Paragraph(dc, "Chapter 1")
                {
                    ParagraphFormat =
                    {
                Style = Heading1Style,
                PageBreakBefore=true
                    }
                });

            // Add SubChapter 1-1.
            section.Blocks.Add(
                        new Paragraph(dc, String.Format("Subchapter 1-1"))
                        {
                            ParagraphFormat =
                            {
                    Style = Heading2Style
                            }
                        });

            // Add the content of Chapter 1 / Subchapter 1-1.
            section.Blocks.Add(
                        new Paragraph(dc,
                            "«Document.Net» will help you in development of applications which operates with DOCX, RTF, PDF, HTML and Text documents.After adding of the reference to(SautinSoft.Document.dll) - it's 100% C# managed assembly you will be able to create a new document, parse an existing, modify anything what you want.")
                        {
                            ParagraphFormat = new ParagraphFormat
                            {
                                LeftIndentation = 10,
                                RightIndentation = 10,
                                SpecialIndentation = 20,
                                LineSpacing = 20,
                                LineSpacingRule = LineSpacingRule.Exactly,
                                SpaceBefore = 20,
                                SpaceAfter = 20
                            }
                        });

            // Let's add another page break into.
            section.Blocks.Add(
                   new Paragraph(dc,
                   new SpecialCharacter(dc, SpecialCharacterType.PageBreak)));

            // Add the document content (Chapter 2).
            // Add Chapter 2.
            section.Blocks.Add(
                     new Paragraph(dc, "Chapter 2")
                     {
                         ParagraphFormat =
                         {
                Style = Heading1Style
                         }
                     });

            // Add SubChapter 2-1.
            section.Blocks.Add(
                new Paragraph(dc, String.Format("Subchapter 2-1"))
                {
                    ParagraphFormat =
                    {
                    Style = Heading2Style
                    }
                });

            // Add the content of Chapter 2 / Subchapter 2-1.
            section.Blocks.Add(
                        new Paragraph(dc,
                            "Requires only .Net 4.0 or above. Our product is compatible with all .Net languages and supports all Operating Systems where .Net Framework can be used. Note that «Document .Net» is entirely written in managed C#, which makes it absolutely standalone and an independent library. Of course, No dependency on Microsoft Word.")
                        {
                            ParagraphFormat = new ParagraphFormat
                            {
                                LeftIndentation = 10,
                                RightIndentation = 10,
                                SpecialIndentation = 20,
                                LineSpacing = 20,
                                LineSpacingRule = LineSpacingRule.Exactly,
                                SpaceBefore = 20,
                                SpaceAfter = 20
                            }
                        });

            // Update TOC (TOC can be updated only after all document content is added).
            var tableofcontents = (TableOfEntries)dc.GetChildElements(true, ElementType.TableOfEntries).FirstOrDefault();
            tableofcontents.Update();

            // Apply the style for the TOC.
            foreach (Paragraph par in tableofcontents.Entries)
            {
                par.ParagraphFormat.Style = TOCStyle;
            }

            // Update TOC's page numbers.
            // Page numbers are automatically updated in that case.
            dc.GetPaginator(new PaginatorOptions() { UpdateFields = true });

            // Save the document as DOCX file.
            dc.Save(resultFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultFile) { UseShellExecute = true });
        }
    }
}