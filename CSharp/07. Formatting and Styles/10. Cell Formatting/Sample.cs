using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using SautinSoft.Document.Tables;
using System.Linq;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            TableCellFormatting();
        }
        /// <summary>
        /// How apply formatting for table cells.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/tablecell-format.php
        /// </remarks>
        static void TableCellFormatting()
        {
            string filePath = @"TableCellFormatting.docx";

            // Let's a new create document.
            DocumentCore dc = new DocumentCore();

            // Add new table.
            Table table = new Table(dc);
            double width = LengthUnitConverter.Convert(100, LengthUnit.Millimeter, LengthUnit.Point);
            table.TableFormat.PreferredWidth = new TableWidth(width, TableWidthUnit.Point);
            table.TableFormat.Borders.SetBorders(MultipleBorderTypes.Inside | MultipleBorderTypes.Outside, 
                BorderStyle.Single, Color.Black, 1);
            dc.Sections.Add(new Section(dc, table));

            TableRow row = new TableRow(dc);
            row.RowFormat.Height = new TableRowHeight(50, HeightRule.Exact);
            table.Rows.Add(row);

            // Create a cells with formatting: borders, background color, vertical alignment and text direction.
            TableCell cell = new TableCell(dc, new Paragraph(dc, "Cell 1"));
            cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Blue, 1);
            cell.CellFormat.BackgroundColor = Color.Yellow;
            row.Cells.Add(cell);

            row.Cells.Add(new TableCell(dc, new Paragraph(dc, "Cell 2"))
            {
                CellFormat = new TableCellFormat()
                {
                    VerticalAlignment = VerticalAlignment.Center
                }
            });

            row.Cells.Add(new TableCell(dc, new Paragraph(dc, "Cell 3"))
            {
                CellFormat = new TableCellFormat()
                {
                    TextDirection = TextDirection.BottomToTop
                }
            });

            // Save our document.
            dc.Save(filePath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}