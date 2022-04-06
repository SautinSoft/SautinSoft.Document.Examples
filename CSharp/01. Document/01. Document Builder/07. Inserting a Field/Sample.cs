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
            InsertingField();
        }
        /// <summary>
        /// Generate document with forms and fields using DocumentBuilder.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-field.php
        /// </remarks>

        static void InsertingField()
        {
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            string resultPath = @"Result.pdf";
            string[] items = { "One", "Two", "Three", "Four", "Five" };

            // Insert the formatted text into the document using DocumentBuilder.
            db.CharacterFormat.FontName = "Verdana";
            db.CharacterFormat.Size = 16;
            db.CharacterFormat.FontColor = Color.Orange;
            db.Writeln("Generate document with forms and fields using DocumentBuilder.");
            db.CharacterFormat.ClearFormatting();

            db.Write(@"{ TIME   \* MERGEFORMAT } - ");
            db.InsertField("TIME");
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);

            db.InsertTextInput("TextInput", FormTextType.RegularText, "", "Insert Text Input", 0);
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);
            
            db.InsertCheckBox("CheckBox", true, 0);
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);

            db.InsertComboBox("DropDown", items, 3);

            // Save our document into PDF format.
            dc.Save(resultPath, new PdfSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}