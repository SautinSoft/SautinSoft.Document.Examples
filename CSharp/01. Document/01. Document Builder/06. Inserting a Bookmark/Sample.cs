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
            InsertingBookmark();
        }
        /// <summary>
        /// How to insert a Bookmark in a document using DocumentBuilder.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-bookmark.php
        /// </remarks>

        static void InsertingBookmark()
        {
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            string resultPath = @"Result.docx";

            // Insert the formatted text into the document using DocumentBuilder.
            db.CharacterFormat.FontName = "Verdana";
            db.CharacterFormat.Size = 16;
            db.CharacterFormat.FontColor = Color.Orange;
            db.Writeln("This text is inserted by the DocumentBuilder.Write method with formatting.");

            // Marks the current position in the document as a 1st bookmark start.
            db.StartBookmark("Firstbookmark");
            db.CharacterFormat.Italic = true;
            db.CharacterFormat.Size = 12;
            db.CharacterFormat.FontColor = Color.Blue;
            db.Writeln("The text inside the bookmark 'Firstbookmark' is inserted by the DocumentBuilder.Writeln method.");
            
            // Marks the current position in the document as a 1st bookmark end.
            db.EndBookmark("Firstbookmark");

            // Insert text after the 1st bookmark.
            db.CharacterFormat.Italic = false;
            db.CharacterFormat.Size = 16;
            db.CharacterFormat.FontColor = Color.Orange;
            db.Writeln("DocumentBuilder.EndBookmark method with the same name points to the end of the bookmark.");

            // Marks the current position in the document as a 2nd bookmark start.
            db.StartBookmark("Secondbookmark");
            db.CharacterFormat.Italic = true;
            db.CharacterFormat.Size = 12;
            db.CharacterFormat.FontColor = Color.Blue;
            db.Writeln("Incorrectly spelled bookmarks or bookmarks with duplicate names will be ignored when saving the document.");

            // Marks the current position in the document as a 2nd bookmark end.
            db.EndBookmark("Secondbookmark");

            // Save our document into DOCX format.
            dc.Save(resultPath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}