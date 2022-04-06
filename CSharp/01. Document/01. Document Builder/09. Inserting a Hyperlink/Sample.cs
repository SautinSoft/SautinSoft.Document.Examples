using System;
using SautinSoft.Document;
using System.Text;
using SautinSoft.Document.Drawing;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertingHyperlink();
        }
        /// <summary>
        /// Insert a hyperlink into a document using DocumentBuilder.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-hyperlink.php
        /// </remarks>

        static void InsertingHyperlink()
        {
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            // Insert the formatted text into the document.
            db.CharacterFormat.FontName = "Verdana";
            db.CharacterFormat.Size = 16;
            db.Writeln("Insert a hyperlink into a document using DocumentBuilder.");

            // Inserts a Word field into a document.
            db.CharacterFormat.Size = 26;
            db.CharacterFormat.FontColor = Color.Brown;
            db.InsertField("DATE");
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);

            // Insert URL hyperlink.
            db.CharacterFormat.FontColor = Color.Blue;
            db.CharacterFormat.UnderlineStyle = UnderlineType.Dashed;
            db.InsertHyperlink("Welcome to SautinSoft!", "https://sautinsoft.com", false);

            db.InsertSpecialCharacter(SpecialCharacterType.PageBreak);

            // Insert a hyperlink inside a document as a bookmark.
            db.CharacterFormat.FontColor = Color.Brown;
            db.CharacterFormat.UnderlineStyle = UnderlineType.DotDotDash;
            db.InsertHyperlink("back to the field {DATE}", "DATE", true);

            // Save our document into DOCX format.
            string resultPath = @"Result.docx";
            dc.Save(resultPath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}