
using System.IO;
using System.Linq;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            MergeRowsAndCellsInTable();
        }

        /// <summary>
        /// Create a new table with rows and cells merged by vertical (rowspan) and horizontal (colspan).
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/table-with-merged-rows-and-cells.php
        /// </remarks>
        public static void MergeRowsAndCellsInTable()
        {

            DocumentCore dc = new DocumentCore();

            Table table = new Table(dc,
                   new TableRow(dc,
                            new TableCell(dc, new Paragraph(dc, "Cell 1-1")),
                            new TableCell(dc, new Paragraph(dc, "Cell 1-2")),
                            new TableCell(dc, new Paragraph(dc, "Cell 1-3")),
                            new TableCell(dc, new Paragraph(dc, "Cell 1-4"))),
                   
                   new TableRow(dc,
                            new TableCell(dc, new Paragraph(dc, "Cell 2-1 -> 3-2"))
                            {
                                RowSpan = 2,
                                ColumnSpan = 2
                            },
                            new TableCell(dc, new Paragraph(dc, "Cell 2-3 -> 2-4"))
                            {
                                ColumnSpan = 2
                            }),
                   
                   new TableRow(dc,
                            new TableCell(dc) { ColumnSpan = 2 },
                            new TableCell(dc, new Paragraph(dc, "Cell 3-3")),
                            new TableCell(dc, new Paragraph(dc, "Cell 3-4"))),
                   
                   new TableRow(dc,
                            new TableCell(dc, new Paragraph(dc, "Cell 4-1"))),
                   new TableRow(dc,
                            new TableCell(dc, new Paragraph(dc, "Cell 5-1"))));

            table.TableFormat.DefaultCellPadding = new Padding(10, LengthUnit.Pixel);

            // Set the table width to 10 cm and convert it to points.
            double tableWidthInPoints = LengthUnitConverter.Convert(10, LengthUnit.Centimeter, LengthUnit.Point);
            table.TableFormat.PreferredWidth = new TableWidth(tableWidthInPoints, TableWidthUnit.Point);
            for (int r = 0; r < table.Rows.Count; r++)
            {
                for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                {
                    TableCell cell = table.Rows[r].Cells[c];
                    cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dashed, Color.Black, 1);
                    cell.CellFormat.BackgroundColor = new Color("#FFCC00");
                }
            }

            dc.Sections.Add(new Section(dc, table));

            dc.Save("MergedTableCells.docx");
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("MergedTableCells.docx") { UseShellExecute = true });
        }
    }
}