using SautinSoft.Document;
using SautinSoft.Document.Tables;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            CreateStyles();
        }
        /// <summary>
        /// This sample shows how to create new styles and apply them to various elements.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/styles-create.php
        /// </remarks>
        public static void CreateStyles()
        {
            string documentPath = @"Styles.docx";

            // Let's create document.
            DocumentCore dc = new DocumentCore();

            // Create 3 styles: CharacterStyle, ParagraphStyle and TableStyle

            // 1. Create CharacterStyle
            CharacterStyle charStyle = new CharacterStyle("Green");
            charStyle.CharacterFormat = new CharacterFormat()
            {
                FontName = "Calibri",
                Size = 20,
                FontColor = Color.Green,
                UnderlineStyle = UnderlineType.Single
            };

            // 2. Create ParagraphStyle
            ParagraphStyle parStyle = new ParagraphStyle("Center");
            parStyle.ParagraphFormat.Alignment = HorizontalAlignment.Center;

            // 3. Create TableStyle
            TableStyle tblStyle = new TableStyle("Blue Table");
            tblStyle.CellFormat.BackgroundColor = new Color("#358CCB");            

            // 4. Add the styles to the style collection.
            dc.Styles.Add(charStyle);
            dc.Styles.Add(parStyle);
            dc.Styles.Add(tblStyle);

            // 5. Add some document content and apply the styles.
            // Add a paragraph and text.

            dc.Sections.Add(
            new Section(dc,
                new Paragraph(dc,
                new Run(dc, "Once upon a time, in a far away swamp, there lived an ogre named "),
                new Run(dc, "Shrek") { CharacterFormat = { Style = charStyle } },
                new Run(dc, " whose precious solitude is suddenly shattered by an invasion of annoying fairy tale characters..."))
                { ParagraphFormat = { Style = parStyle } }));

            // Add a table (1 x 2).
            Table tbl = new Table(dc, 1, 2);
            tbl.TableFormat.Style = tblStyle;
            tbl.Rows[0].Cells[0].Content.Start.Insert("Column 1");
            tbl.Rows[0].Cells[1].Content.Start.Insert("Column 2");
            dc.Sections[0].Blocks.Add(tbl);

            // Save our document into DOCX format.
            dc.Save(documentPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}