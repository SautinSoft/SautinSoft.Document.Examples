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
            TableRowFormatting();
        }
        /// <summary>
        /// Shows how to set a height for a table row, repeat a row as header on each page, shift a row by N columns to the right.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/tablerow-format.php
        /// </remarks>
        static void TableRowFormatting()
        {
            string docxPath = @"FormattedTable.docx";

            // Let's create document.
            DocumentCore dc = new DocumentCore();

            Section s = new Section(dc);
            dc.Sections.Add(s);


            int rows = 30;
            int columns = 5;

            Table t = new Table(dc, rows, columns);
            t.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);
            t.TableFormat.Borders.SetBorders(MultipleBorderTypes.All, BorderStyle.Single, Color.DarkGray, 1);
            t.TableFormat.AutomaticallyResizeToFitContents = false;
            s.Blocks.Add(t);

            // Specify row height:
            // 10 mm - for odd rows.
            // 15 mm - for even rows.
            double oddHeight = LengthUnitConverter.Convert(10, LengthUnit.Millimeter, LengthUnit.Point);
            double evenHeight = LengthUnitConverter.Convert(15, LengthUnit.Millimeter, LengthUnit.Point);
            for (int r = 0; r < t.Rows.Count; r++)
            {
                TableRow row = t.Rows[r];
                if (r % 2 != 0)
                    row.RowFormat.Height = new TableRowHeight(evenHeight, HeightRule.AtLeast);
                else
                    row.RowFormat.Height = new TableRowHeight(oddHeight, HeightRule.AtLeast);
            }

            // Add the table caption - mark the specific row (for example: 0) to repeat on each page.
            TableRow firstRow = t.Rows[0];
            // Repeate as header row at the top of each page.
            // Note: Only the first row in the table can be set up as header.
            firstRow.RowFormat.RepeatOnEachPage = true;

            // Merge all cells into a one in the first row (Caption).
            int colSpan = firstRow.Cells.Count;
            for (int c = firstRow.Cells.Count - 1; c>=1; c--)
            {
                firstRow.Cells.RemoveAt(c);
            }
            // Specify how many columns this cell will take up.
            firstRow.Cells[0].ColumnSpan = colSpan;

            // Set the table caption in the first row and first cell. 
            Paragraph p = new Paragraph(dc);
            p.Inlines.Add(new Run(dc, "This is the Row 0 (RepeatOnEachPage = true)", new CharacterFormat() { FontColor = Color.Blue, Size = 20 }));
            p.ParagraphFormat.Alignment = HorizontalAlignment.Center;
            t.Rows[0].Cells[0].Blocks.Add(p);

            // Another interesting properties of TableRowFormat:
            // GridBefore and GridAfter
            // Add "Total" at the end of the table.
            TableRow rowTotal = new TableRow(dc);
            rowTotal.Cells.Add(new TableCell(dc));
            rowTotal.Cells[0].Content.Start.Insert(string.Format("Total rows: {0}", rows), new CharacterFormat() { FontColor = Color.Red, Size = 30 });
            
            // Shift the rowTotal to the right corner.
            // In our case, shift on 4 columns.
            rowTotal.RowFormat.GridBefore = columns-1;
            

            rowTotal.RowFormat.Height = new TableRowHeight(evenHeight, HeightRule.AtLeast);
            t.Rows.Add(rowTotal);            

            // Save our document into DOCX format.
            dc.Save(docxPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docxPath) { UseShellExecute = true });
        }
    }
}