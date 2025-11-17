using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            // "Right to left" (RTL) refers to the direction in which text is written and read, starting from the right side of the page or line and moving to the left.
            // This writing direction is used in languages such as Arabic, Hebrew, Persian, and Urdu.
            // It also influences the layout of user interfaces and documents in these languages

            Create_WORD_RTL();

        }
        /// <summary>
        /// Right-to-Left. Converting Word file to PDF without losing any formatting.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-RTL-word-document-net-csharp-vb.php
        /// </remarks>
        public static void Create_WORD_RTL()
        {
            // Set a path to our document.
            string docPath = @"Right-to-Left.docx";

            // Create a new document and DocumentBuilder.
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            // Set page size A4.
            Section section = db.Document.Sections[0];
            section.PageSetup.PaperType = PaperType.A4;

            // Add 1st paragraph with formatted text.
            db.CharacterFormat.FontName = "Verdana";
            db.CharacterFormat.Size = 16;
            db.CharacterFormat.FontColor = Color.Orange;
            db.Write("أخذ عن موالية الإمتعاض");
            // Add a line break into the 1st paragraph.
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);
            // Add 2nd line to the 1st paragraph, create 2nd paragraph.
            db.Writeln("של תיבת תרומה מלא");
            // Specify the paragraph alignment.
            (section.Blocks[0] as Paragraph).ParagraphFormat.Alignment = HorizontalAlignment.Center;

            // Add text into the 2nd paragraph.
            db.CharacterFormat.ClearFormatting();
            db.CharacterFormat.Size = 25;
            db.CharacterFormat.FontColor = Color.Blue;
            db.CharacterFormat.Bold = true;
            db.Write("اُردُو حُرُوفِ تَہَجِّی‌");
            // Insert a line break into the 2nd paragraph.
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);
            // Insert 2nd line with own formatting to the 2nd paragraph.
            db.CharacterFormat.Size = 20;
            db.CharacterFormat.FontColor = Color.DarkGreen;
            db.CharacterFormat.UnderlineStyle = UnderlineType.Single;
            db.CharacterFormat.Bold = false;
            db.Write("هلو می فریند. ");

            // Add a graphics figure into the paragraph.
            db.CharacterFormat.ClearFormatting();
            Shape shape = db.InsertShape(SautinSoft.Document.Drawing.Figure.SmileyFace, new SautinSoft.Document.Drawing.Size(50, 50, LengthUnit.Millimeter));
            // Specify outline and fill.
            shape.Outline.Fill.SetSolid(new SautinSoft.Document.Color(53, 140, 203));
            shape.Outline.Width = 3;
            shape.Fill.SetSolid(SautinSoft.Document.Color.White);

            // Save the document to the file in DOCX format.
            dc.Save(docPath, new DocxSaveOptions());

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docPath) { UseShellExecute = true });
        }
    }
}