using System;
using SautinSoft.Document;
using System.Text;


namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertingParagraph();
        }
        /// <summary>
        /// Create a document and insert a paragraph using DocumentBuilder.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-paragraph.php
        /// </remarks>

        static void InsertingParagraph()
        {
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            string resultPath = @"Result.docx";

            // Insert the formatted text into the document using DocumentBuilder.
            db.CharacterFormat.FontName = "Verdana";
            db.CharacterFormat.Size = 16.5f;
            db.CharacterFormat.AllCaps = true;
            db.CharacterFormat.Italic = true;
            db.CharacterFormat.FontColor = Color.Orange;
            db.ParagraphFormat.LeftIndentation = 30;
            db.Writeln("This paragraph has a Left Indentation of 30 points.");
            db.ParagraphFormat.SpecialIndentation = 50;
            db.Writeln("This paragraph retains the Left Indentation of 30 points and is supplemented by the first-line indent of 50 points.");

            // This method will clear all directly set formatting values.
            db.ParagraphFormat.ClearFormatting();
            db.CharacterFormat.ClearFormatting();
            db.Write("All directly set text and paragraph formatting values were cleared using DocumentBuilder.");

            // Save the document to the file in DOCX format.
            dc.Save(resultPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}