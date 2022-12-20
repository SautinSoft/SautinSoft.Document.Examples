using System;
using System.IO;
using System.Collections.Generic;
using SautinSoft.Document;
using System.Linq;
using SautinSoft.Document.Tables;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            TableWithDocumentCore();
            TableWithDocumentBuilder();
        }

        /// <summary>
        /// This sample shows how to creating a custom table with formatting using DocumentCore or DocumentBuilder classes.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-creating-custom-table-with-formatting-in-csharp-vb-net.php
        /// </remarks>
        public static void TableWithDocumentCore()
        {
            string documentPath = @"tableDC.pdf";

            // Let's create a new document.
            DocumentCore dc = new DocumentCore();

            Section s = new Section(dc);

            TableFormat tf = new TableFormat();
            tf.Borders.ClearBorders();
            tf.AutomaticallyResizeToFitContents = false;

            var table = new Table(dc);
           
            // Add columns with specified width.
            table.Columns.Add(new TableColumn(60));
            table.Columns.Add(new TableColumn(120));
            table.Columns.Add(new TableColumn(180));

            // Add rows with specified height.
            table.Rows.Add(new TableRow(dc));
            table.Rows[0].RowFormat.Height = new TableRowHeight(30, HeightRule.AtLeast);
            table.Rows.Add(new TableRow(dc));
            table.Rows[1].RowFormat.Height = new TableRowHeight(60, HeightRule.AtLeast);
            table.Rows.Add(new TableRow(dc));
            table.Rows[2].RowFormat.Height = new TableRowHeight(90, HeightRule.AtLeast);

            for (int r = 0; r < 3; r++)
                for (int c = 0; c < 3; c++)
                {
                    // Add cell.
                    var cell = new TableCell(dc);
                    table.Rows[r].Cells.Add(cell);

                    // Set cell's vertical alignment.
                    cell.CellFormat.VerticalAlignment = (VerticalAlignment)r;

                    // Add cell content.
                    var paragraph = new Paragraph(dc, $"Cell ({r + 1},{c + 1})");
                    cell.Blocks.Add(paragraph);

                    // Set cell content's horizontal alignment.
                    paragraph.ParagraphFormat.Alignment = (HorizontalAlignment)c;

                    if ((r + c) % 2 == 0)
                    {
                        // Set cell's background and borders.
                        cell.CellFormat.BackgroundColor = new Color(255, 242, 204);
                        cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Double, Color.Red, 1);
                    }
                }

            dc.Sections.Add(new Section(dc, table));

            // Save our document into PDF format.
            dc.Save(documentPath, new PdfSaveOptions() { Compliance = PdfCompliance.PDF_A1a });

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
		
        public static void TableWithDocumentBuilder()
        {
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            // Create a new table with preferred width.
            Table table = db.StartTable();
			
           // db.TableFormat.PreferredWidth = new TableWidth(LengthUnitConverter.Convert(5, LengthUnit.Inch, LengthUnit.Point), TableWidthUnit.Point);
            db.TableFormat.AutomaticallyResizeToFitContents = false;
			
            // Specify formatting of cells and alignment.
            db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Green, 1);
            db.CellFormat.VerticalAlignment = VerticalAlignment.Top;
            table.Columns.Add(new TableColumn() { PreferredWidth = 100 });
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
            table.Columns.Add(new TableColumn() { PreferredWidth = 100 });
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
            table.Columns.Add(new TableColumn() { PreferredWidth = 150 });
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
            string filePath = "tableDB.docx";
            dc.Save(filePath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
    }
   }
 }