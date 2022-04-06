using System;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            TableFormat();
        }

		/// <summary>
        /// How to create a table and apply formatting in a document. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/table-format.php
        /// </remarks>		
        public static void TableFormat()
        {
            string documentPath = @"TableFormat.docx";

            // Let's create a new document.
            DocumentCore dc = new DocumentCore();

            // Add a new section.
            Section s = new Section(dc);
            dc.Sections.Add(s);

            Table table = new Table(dc);
            table.TableFormat.AutomaticallyResizeToFitContents = false;
            table.TableFormat.Alignment = HorizontalAlignment.Center;
            table.TableFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dotted, Color.DarkBlue, 2.0);

            table.Columns.Add(new TableColumn() { PreferredWidth = 50 });
            table.Columns.Add(new TableColumn() { PreferredWidth = 80 });
            table.Columns.Add(new TableColumn() { PreferredWidth = 100 });
            table.Columns.Add(new TableColumn() { PreferredWidth = 120 });
            table.Columns.Add(new TableColumn() { PreferredWidth = 80 });

            TableRow row = new TableRow(dc);
            row.RowFormat.Height = new TableRowHeight(100, HeightRule.AtLeast);
            table.Rows.Add(row);

            TableCell cell1 = new TableCell(dc, new Paragraph(dc, "PDF is Portable Document Format"));
            cell1.CellFormat.TextDirection = TextDirection.TopToBottom;
            cell1.CellFormat.BackgroundColor = Color.Yellow;
            cell1.CellFormat.Padding = new Padding(5);
            cell1.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dotted, Color.Yellow, 1.0);
            row.Cells.Add(cell1);

            TableCell cell2 = new TableCell(dc, new Paragraph(dc, "DOCX is Office Open XML"));
            cell2.CellFormat.VerticalAlignment = VerticalAlignment.Center;
            cell2.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dotted, Color.Black, 1.0);
            row.Cells.Add(cell2);

            row.Cells.Add(new TableCell(dc, new Paragraph(dc, "HTML is Hypertext Markup Language"))
            {
                CellFormat = new TableCellFormat()
                {
                    BackgroundColor = Color.Pink                   
                }
            });

            row.Cells.Add(new TableCell(dc, new Paragraph(dc, "Images: jpeg, png, bmp, tiff")
            {
                ParagraphFormat = new ParagraphFormat()
                {
                    Alignment = HorizontalAlignment.Center
                }
            })
            {
                CellFormat = new TableCellFormat()
                {
                    VerticalAlignment = VerticalAlignment.Center,
                    BackgroundColor = Color.Purple
                }
            });

            TableCell cell5 = new TableCell(dc, new Paragraph(dc, "RTF is Rich Text Format"));
            cell5.CellFormat.TextDirection = TextDirection.BottomToTop;
            cell5.CellFormat.Padding = new Padding(5);
            cell5.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dashed, Color.Red, 1.0);
            row.Cells.Add(cell5);
            
            // Add the table into the section.
            s.Blocks.Add(table);

            // Save our document DOCX.
            dc.Save(documentPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}