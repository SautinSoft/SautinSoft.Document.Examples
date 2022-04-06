using System;
using System.Text;
using SautinSoft.Document;
using SautinSoft.Document.CustomMarkups;
using System.IO;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertPlainText();
        }
        /// <summary>
        /// Inserting a plain text content control.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/content-controls-insert-plain-text-net-csharp-vb.php
        /// </remarks>

        static void InsertPlainText()
        {
            // Let's create a simple document.
            DocumentCore dc = new DocumentCore();

            // Create a plain text content control.
            BlockContentControl pt = new BlockContentControl(dc, ContentControlType.PlainText);

            // Add a new section.
            dc.Sections.Add(new Section(dc, pt));

            // Add the content control properties.
            pt.Properties.Title = "Title";
            pt.Properties.Multiline = true;
            pt.Properties.Color = Color.Blue;
            pt.Document.DefaultCharacterFormat.FontColor = Color.Orange;

            // Add new paragraph with formatted text.
            pt.Blocks.Add(new Paragraph(dc,
            new Run(dc, "This is first paragraph with symbols added on a new line."),
            new SpecialCharacter(dc, SpecialCharacterType.LineBreak),
            new Run(dc, "This is a new line in the first paragraph."),
            new SpecialCharacter(dc, SpecialCharacterType.LineBreak),
            new Run(dc, "Insert the \"Wingdings\" font family with formatting."),
            new Run(dc, "\xFC" + "\xF0" + "\x32") { CharacterFormat = { FontName = "Wingdings", FontColor = new Color("#000000"), Size = 48 }}));

            // Save our document into DOCX format.
            string resultPath = @"result.docx";
            dc.Save(resultPath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}
