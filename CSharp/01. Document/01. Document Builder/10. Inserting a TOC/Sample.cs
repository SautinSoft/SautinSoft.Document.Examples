using System;
using SautinSoft.Document;
using System.Text;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertingToc();
        }
        /// <summary>
        /// Insert a TOC (Table of Contents) field into the document using DocumentBuilder.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-toc.php
        /// </remarks>

        static void InsertingToc()
        {
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            // Insert the formatted text into the document using DocumentBuilder.
            db.CharacterFormat.FontName = "Verdana";
            db.CharacterFormat.Size = 16;
            db.Writeln("Table of Contents.");
            db.CharacterFormat.ClearFormatting();

            // Insert Table of Contents field into the document at the current position.
            TableOfEntries toe = db.InsertTableOfContents("\\o \"1-3\" \\h");
            // For information about switches, see the description on the page above.

            // Add the text and divide it into headings.
            db.InsertSpecialCharacter(SpecialCharacterType.PageBreak);
            ParagraphStyle Heading1Style = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Heading1, dc);
            dc.Styles.Add(Heading1Style);
            db.ParagraphFormat.Style = Heading1Style;
            db.Writeln("Heading 1");
            db.ParagraphFormat.ClearFormatting();
            db.Writeln("Some text Heading 1 Some text Heading 1 Some text Heading 1 Some text Heading 1 Some text Heading 1 Some text Heading 1" +
                "Some text Heading 1 Some text Heading 1 Some text Heading 1 Some text Heading 1 Some text Heading 1 Some text Heading 1 ");

            ParagraphStyle Heading2Style = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Heading2, dc);
            dc.Styles.Add(Heading2Style);
            db.ParagraphFormat.Style = Heading2Style;
            db.Writeln("Heading 1.1");
            db.ParagraphFormat.ClearFormatting();
            db.Writeln("Some text Heading 1.1 Some text Heading 1.1 Some text Heading 1.1 Some text Heading 1.1 Some text Heading 1.1" +
                " Some text Heading 1.1 Some text Heading 1.1 Some text Heading 1.1 Some text Heading 1.1");
            db.ParagraphFormat.Style = Heading2Style;
            db.Writeln("Heading 1.2");
            db.ParagraphFormat.ClearFormatting();
            db.Writeln("Some text Heading 1.2 Some text Heading 1.2 Some text Heading 1.2 Some text Heading 1.2 Some text Heading 1.2" +
                " Some text Heading 1.2 Some text Heading 1.2 Some text Heading 1.2 ");

            ParagraphStyle Heading3Style = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Heading3, dc);
            dc.Styles.Add(Heading3Style);
            db.ParagraphFormat.Style = Heading3Style;
            db.Writeln("Heading 1.1.1");
            db.ParagraphFormat.ClearFormatting();
            db.Writeln("Some text Heading 1.1.1 Some text Heading 1.1.1 Some text Heading 1.1.1 Some text Heading 1.1.1 " +
                " Some text Heading 1.1.1 Some text Heading 1.1.1 Some text Heading 1.1.1 ");
            db.ParagraphFormat.Style = Heading3Style;
            db.Writeln("Heading 1.1.2");
            db.ParagraphFormat.ClearFormatting();
            db.Writeln("Some text 1.1.2 Some text 1.1.2 Some text 1.1.2 Some text 1.1.2");

            db.ParagraphFormat.Style = Heading1Style;
            db.Writeln("Heading 2");
            db.ParagraphFormat.ClearFormatting();
            db.Writeln("Some text Heading 2 Some text Heading 2.");

            db.ParagraphFormat.Style = Heading1Style;
            db.Writeln("Heading 3");
            db.ParagraphFormat.ClearFormatting();
            db.Writeln("Some text Heading 3 Some text Heading 3 Some text Heading 3 Some text Heading 3 Some text Heading 3" +
                 "Some text Heading 3Some text Heading 3Some text Heading 3Some text Heading 3Some text Heading 3");

            db.ParagraphFormat.Style = Heading2Style;
            db.Writeln("Heading 3.1");
            db.ParagraphFormat.ClearFormatting();
            db.Writeln("Some text Heading 3.1 Some text Heading 3.1 Some text Heading 3.1 Some text Heading 3.1 Some text Heading 3.1" +
               "Some text Heading 3.1 Some text Heading 3.1 Some text Heading 3.1");
            
            db.ParagraphFormat.Style = Heading2Style;
            db.Writeln("Heading 3.2");
            db.ParagraphFormat.ClearFormatting();
            db.Writeln("Some text Heading 3.2 Some text Heading 3.2 Some text Heading 3.2 Some text Heading 3.2 Some text Heading 3.2");

            // Update the TOC field (table of contents).
            toe.Update();

            // Save our document into DOCX format.
            string resultPath = @"Result.docx";
            dc.Save(resultPath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}