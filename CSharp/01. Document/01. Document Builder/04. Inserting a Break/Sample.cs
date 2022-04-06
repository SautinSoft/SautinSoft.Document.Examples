using System;
using SautinSoft.Document;
using System.Text;


namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertingBreak();
        }
        /// <summary>
        /// Insert a Line Break, Column Break, Page Break using DocumentBuilder.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-break.php
        /// </remarks>

        static void InsertingBreak()
        {
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            string resultPath = @"Result.docx";
            db.PageSetup.TextColumns = new TextColumnCollection(2);

            // Insert the formatted text into the document using DocumentBuilder.
            db.CharacterFormat.FontName = "Verdana";
            db.CharacterFormat.Size = 16.5f;
            db.CharacterFormat.AllCaps = true;
            db.CharacterFormat.Italic = true;
            db.CharacterFormat.FontColor = Color.Orange;
            db.ParagraphFormat.LeftIndentation = 30;
            db.Writeln("This paragraph has a Left Indentation of 30 points.");

            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);

            // Undo the previously applied formatting.
            db.ParagraphFormat.ClearFormatting();
            db.CharacterFormat.ClearFormatting();

            db.Writeln("After this paragraph insert a column break.");
            db.InsertSpecialCharacter(SpecialCharacterType.ColumnBreak);

            db.CharacterFormat.Italic = true;
            db.CharacterFormat.FontColor = Color.DarkBlue;
            db.CharacterFormat.Size = 20f;
            db.Writeln("After this paragraph insert a page break.");
            db.InsertSpecialCharacter(SpecialCharacterType.PageBreak);

            // Save the document to the file in DOCX format.
            dc.Save(resultPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}