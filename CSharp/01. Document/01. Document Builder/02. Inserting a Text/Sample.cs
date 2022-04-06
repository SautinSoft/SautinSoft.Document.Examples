using System;
using SautinSoft.Document;
using System.Text;


namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertingText();
        }
        /// <summary>
        /// Create a document and insert a string of text using DocumentBuilder.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-text.php
        /// </remarks>

        static void InsertingText()
        {
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            string resultPath = @"Result.pdf";

            // Insert the formatted text into the document using DocumentBuilder.
            db.CharacterFormat.FontName = "Verdana";
            db.CharacterFormat.Size = 55.5f;
            db.CharacterFormat.AllCaps = true;
            db.CharacterFormat.Italic = true;
            db.CharacterFormat.FontColor = Color.Orange;
            db.Write("insert a text using");

            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);

            db.CharacterFormat.Size = 52.5f;
            db.CharacterFormat.FontColor = Color.Blue;
            db.CharacterFormat.AllCaps = false;
            db.CharacterFormat.Italic = false;
            db.Write("DocumentBuilder");

            // Save the document to the file in PDF format.
            dc.Save(resultPath, new PdfSaveOptions()
            { Compliance = PdfCompliance.PDF_A1a });

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}