using System;
using System.IO;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            SaveToXmlFile();
        }

        /// <summary>
        /// Creates a new document and saves it as Xml file.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/save-document-as-xml-net-csharp-vb.php
        /// </remarks>
        static void SaveToXmlFile()
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

            // Save our document into XML format.
            string filePath = "Result.xml";
            dc.Save(filePath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}