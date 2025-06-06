using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using SkiaSharp;
using System.Drawing;
using System.IO;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            // Here we'll show two ways to create Image document from a scratch.
            // Use any of them, which is more clear to you.

            // 1. With help of DocumentBuilder (wizard).
            CreateImageUsingDocumentBuilder();

            // 2. With Document Object Model (DOM) directly.
            CreateImageUsingDOM();
        }
        /// <summary>
        /// Creates a new Image document using DocumentBuilder wizard.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-image-document-net-csharp-vb.php
        /// </remarks>
        public static void CreateImageUsingDocumentBuilder()
        {
            // Set a path to our document.
            string docPath = @"Result-DocumentBuilder.png";

            // Create a new document and DocumentBuilder.
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            // Set page size A4.
            Section section = db.Document.Sections[0];
            section.PageSetup.PaperType = PaperType.A4;

            // Add 1st paragraph with formatted text.
            db.CharacterFormat.FontName = "Verdana";
            db.CharacterFormat.Size = 16;
            db.CharacterFormat.FontColor = SautinSoft.Document.Color.Orange;
            db.Write("This is a first line in 1st paragraph!");
            // Add a line break into the 1st paragraph.
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);
            // Add 2nd line to the 1st paragraph, create 2nd paragraph.
            db.Writeln("Let's type a second line.");
            // Specify the paragraph alignment.
            (section.Blocks[0] as Paragraph).ParagraphFormat.Alignment = HorizontalAlignment.Center;

            // Add text into the 2nd paragraph.
            db.CharacterFormat.ClearFormatting();
            db.CharacterFormat.Size = 25;
            db.CharacterFormat.FontColor = SautinSoft.Document.Color.Blue;
            db.CharacterFormat.Bold = true;
            db.Write("This is a first line in 2nd paragraph.");
            // Insert a line break into the 2nd paragraph.
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);
            // Insert 2nd line with own formatting to the 2nd paragraph.
            db.CharacterFormat.Size = 20;
            db.CharacterFormat.FontColor = SautinSoft.Document.Color.DarkGreen;
            db.CharacterFormat.UnderlineStyle = UnderlineType.Single;
            db.CharacterFormat.Bold = false;
            db.Write("This is a second line.");

            // Add a graphics figure into the paragraph.
            db.CharacterFormat.ClearFormatting();
            Shape shape = db.InsertShape(SautinSoft.Document.Drawing.Figure.SmileyFace, new SautinSoft.Document.Drawing.Size(50, 50, LengthUnit.Millimeter));
            // Specify outline and fill.
            shape.Outline.Fill.SetSolid(new SautinSoft.Document.Color(53, 140, 203));
            shape.Outline.Width = 3;
            shape.Fill.SetSolid(SautinSoft.Document.Color.Orange);

            // Save the 1st document page to the file in PNG format.
            ImageSaveOptions options = new ImageSaveOptions();
            options.DpiX = 300;
            options.DpiY = 300;
            dc.GetPaginator().Pages[0].Save(docPath, options);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docPath) { UseShellExecute = true });
        }


        /// <summary>
        /// Creates a new Image document using DOM directly.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-image-document-net-csharp-vb.php
        /// </remarks>
        public static void CreateImageUsingDOM()
        {
            // Set a path to our document.
            string docPath = @"Result-DocumentCore.png";

            // Create a new document.
            DocumentCore dc = new DocumentCore();

            // Add new section.
            Section section = new Section(dc);
            dc.Sections.Add(section);

            // Let's set page size A4.
            section.PageSetup.PaperType = PaperType.A4;

            // Add two paragraphs            
            Paragraph par1 = new Paragraph(dc);
            par1.ParagraphFormat.Alignment = HorizontalAlignment.Center;
            section.Blocks.Add(par1);

            // Let's create a characterformat for text in the 1st paragraph.
            CharacterFormat cf = new CharacterFormat() { FontName = "Verdana", Size = 16, FontColor = SautinSoft.Document.Color.Orange };
            Run run1 = new Run(dc, "This is a first line in 1st paragraph!");
            run1.CharacterFormat = cf;
            par1.Inlines.Add(run1);

            // Let's add a line break into the 1st paragraph.
            par1.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.LineBreak));
            // Copy the formatting.
            Run run2 = run1.Clone();
            run2.Text = "Let's type a second line.";
            par1.Inlines.Add(run2);

            // Add 2nd paragraph.
            Paragraph par2 = new Paragraph(dc, new Run(dc, "This is a first line in 2nd paragraph.", new CharacterFormat() { Size = 25, FontColor = SautinSoft.Document.Color.Blue, Bold = true }));
            section.Blocks.Add(par2);
            SpecialCharacter lBr = new SpecialCharacter(dc, SpecialCharacterType.LineBreak);
            par2.Inlines.Add(lBr);
            Run run3 = new Run(dc, "This is a second line.", new CharacterFormat() { Size = 20, FontColor = SautinSoft.Document.Color.DarkGreen, UnderlineStyle = UnderlineType.Single });
            par2.Inlines.Add(run3);

            // Add a graphics figure into the paragraph.
            Shape shape = new Shape(dc, new InlineLayout(new SautinSoft.Document.Drawing.Size(50, 50, LengthUnit.Millimeter)));
            // Specify outline and fill.
            shape.Outline.Fill.SetSolid(new SautinSoft.Document.Color(53, 140, 203));
            shape.Outline.Width = 3;
            shape.Fill.SetSolid(SautinSoft.Document.Color.Orange);
            shape.Geometry.SetPreset(Figure.SmileyFace);
            par2.Inlines.Add(shape);

            // Save the 1st document page to the file in PNG format.
            ImageSaveOptions options = new ImageSaveOptions();
            options.DpiX = 300;
            options.DpiY = 300;
            dc.GetPaginator().Pages[0].Save(docPath, options);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docPath) { UseShellExecute = true });
        }
    }
}