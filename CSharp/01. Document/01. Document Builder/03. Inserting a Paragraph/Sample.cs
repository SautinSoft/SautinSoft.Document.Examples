using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free 30-day key here:   
            // https://sautinsoft.com/start-for-free/
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
            // Create a new document.
            DocumentCore dc = new DocumentCore();
            // Initialize DocumentBuilder with the created document.
            DocumentBuilder db = new DocumentBuilder(dc);

            // Insert a new paragraph.
            db.Writeln("This is an example of a paragraph inserted using DocumentBuilder.");
            db.CharacterFormat.FontName = "Verdana";
            db.CharacterFormat.Size = 16.5f;
            db.CharacterFormat.AllCaps = true;
            db.CharacterFormat.Italic = true;
            db.CharacterFormat.FontColor = Color.Orange;
            db.ParagraphFormat.LeftIndentation = 30;
            db.Writeln("This paragraph has a Left Indentation of 30 points.");
            db.ParagraphFormat.SpecialIndentation = 50;
            db.Writeln("This paragraph retains the Left Indentation of 30 points and is supplemented by the first-line indent of 50 points.");

            // Save the document.
            dc.Save("Example.docx");

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("Example.docx") { UseShellExecute = true });
        }
    }
}