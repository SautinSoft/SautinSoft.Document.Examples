using System;
using SautinSoft.Document;
using SautinSoft.Document.Tables;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertingTable();
        }
        /// <summary>
        /// Creates a table using DocumentBuilder and saves it in a desired format.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-table.php
        /// </remarks>

        static void InsertingTable()
        {
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            // Create a new table with preferred width.
            Table table = db.StartTable();
            db.TableFormat.PreferredWidth = new TableWidth(LengthUnitConverter.Convert(5, LengthUnit.Inch, LengthUnit.Point), TableWidthUnit.Point);

            // Specify formatting of cells and alignment.
            db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Green, 1);
            db.CellFormat.VerticalAlignment = VerticalAlignment.Top;
            db.ParagraphFormat.Alignment = HorizontalAlignment.Center;

            // Specify height of rows and write text.
            db.RowFormat.Height = new TableRowHeight(105f, HeightRule.Exact);
            db.InsertCell();
            db.Write("This is Row 1 Cell 1");
            db.InsertCell();
            db.Write("This is Row 1 Cell 2");
            db.InsertCell();
            db.Write("This is Row 1 Cell 3");
            db.EndRow();

            // Specify formatting of cells and alignment.
            db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 1);
            db.CellFormat.VerticalAlignment = VerticalAlignment.Center;
            db.ParagraphFormat.Alignment = HorizontalAlignment.Left;

            // Specify height of rows and write text.
            db.RowFormat.Height = new TableRowHeight(150f, HeightRule.Exact);
            db.InsertCell();
            db.Write("This is Row 2 Cell 1");
            db.InsertCell();
            db.Write("This is Row 2 Cell 2");
            db.InsertCell();
            db.Write("This is Row 2 Cell 3");
            db.EndRow();

            // Specify formatting of cells and alignment.
            db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Orange, 1);
            db.CellFormat.VerticalAlignment = VerticalAlignment.Bottom;
            db.ParagraphFormat.Alignment = HorizontalAlignment.Right;

            // Specify height of rows and write text
            db.RowFormat.Height = new TableRowHeight(125f, HeightRule.Exact);
            db.InsertCell();
            db.Write("This is Row 3 Cell 1");
            db.InsertCell();
            db.Write("This is Row 3 Cell 2");
            db.InsertCell();
            db.Write("This is Row 3 Cell 3");
            db.EndRow();
            db.EndTable();

            // Save our document into DOCX format.
            string filePath = "Result.docx";
            dc.Save(filePath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}