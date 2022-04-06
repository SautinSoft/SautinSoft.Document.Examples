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
            MovingCursor();
        }
        /// <summary>
        /// Moving the current cursor position in the document using DocumentBuilder.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-moving-cursor.php
        /// </remarks>

        static void MovingCursor()
        {
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            db.MoveToHeaderFooter(HeaderFooterType.HeaderDefault);
            db.Writeln("Moved the cursor to the header and inserted this text.");

            db.MoveToDocumentStart();
            db.CharacterFormat.Size = 16;
            db.CharacterFormat.FontColor = Color.Blue;
            db.Writeln("Moved the cursor to the start of the document.");

            // Marks the current position in the document as a 1st bookmark start.
            db.StartBookmark("Firstbookmark");
            db.CharacterFormat.Italic = true;
            db.CharacterFormat.Size = 14;
            db.CharacterFormat.FontColor = Color.Red;
            db.Writeln("The text inside the 'Bookmark' is inserted by the DocumentBuilder.Writeln method.");
            // Marks the current position in the document as a 1st bookmark end.
            db.EndBookmark("Firstbookmark");

            db.MoveToBookmark("Firstbookmark", true, false);
            db.Writeln("Moved the cursor to the start of the Bookmark.");

            db.CharacterFormat.FontColor = Color.Black;
            Field f1 = db.InsertField("DATE");
            db.MoveToField(f1, false);
            db.Write("Before the field");

            // Moving to the Header and insert the table with three cells.
            db.MoveToHeaderFooter(HeaderFooterType.HeaderDefault);
            db.StartTable();
            db.TableFormat.PreferredWidth = new TableWidth(LengthUnitConverter.Convert(6, LengthUnit.Inch, LengthUnit.Point), TableWidthUnit.Point);
            db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Green, 1);
            db.RowFormat.Height = new TableRowHeight(40, HeightRule.Exact);
            db.CharacterFormat.FontColor = Color.Green;
            db.CharacterFormat.Italic = false;
            db.InsertCell();
            db.Write("This is Row 1 Cell 1");
            db.InsertCell();
            db.Write("This is Row 1 Cell 2");
            db.InsertCell();
            db.Write("This is Row 1 Cell 3");
            db.EndTable();

            // Insert the text in the second cell in the sixth position.
            db.MoveToCell(0, 0, 1, 5);
            db.CharacterFormat.Size = 18;
            db.CharacterFormat.FontColor = Color.Orange;
            db.Write("InsertToCell");

            db.MoveToDocumentEnd();
            db.CharacterFormat.Size = 16;
            db.CharacterFormat.FontColor = Color.Blue;
            db.Writeln("Moved the cursor to the end of the document.");

            // Save our document into DOCX format.
            string resultPath = @"Result.docx";
            dc.Save(resultPath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}