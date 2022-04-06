using System;
using SautinSoft.Document;
using System.Text;
using SautinSoft.Document.Drawing;
using SautinSoft.Document.Tables;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            FormattingOnStack();
        }
        /// <summary>
        /// Saves and retrieves current character formatting on the stack.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/character-formatting-stack.php
        /// </remarks>

        static void FormattingOnStack()
        {
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            db.CharacterFormat.FontName = "Arial";
            db.CharacterFormat.Size = 16;
            db.CharacterFormat.FontColor = Color.Blue;
            db.Writeln("This text contains formatting font name, size and color. Save the character formatting of this text in as first element of the stack.");
            db.PushCharacterFormat();
            db.CharacterFormat.ClearFormatting();

            db.CharacterFormat.Size = 26;
            db.CharacterFormat.FontColor = Color.Orange;
            db.CharacterFormat.Italic = true;
            db.Writeln("This text contains formatting for font size, color and italic. Save the character formatting of this text as second element of the stack.");
            db.PushCharacterFormat();
            db.CharacterFormat.ClearFormatting();

            // Insert the third way the character formatting of the text.
            db.CharacterFormat.Size = 14;
            db.CharacterFormat.FontColor = Color.Red;
            db.CharacterFormat.Bold = true;
            db.Writeln("This text contains formatting for font size, color and bold.");
            db.CharacterFormat.ClearFormatting();

            // Retrieves text character formatting from the stack (the second element).
            db.PopCharacterFormat();

            // Retrieves text character formatting from the stack (the first element).
            db.PopCharacterFormat();
            db.Writeln("The character formatting of this text is extracted from the stack as first element.");

            // Save our document into DOCX format.
            string resultPath = @"Result.docx";
            dc.Save(resultPath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}