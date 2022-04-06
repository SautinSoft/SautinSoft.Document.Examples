using SautinSoft.Document;
using SautinSoft.Document.Tables;
using System;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            InsertTableAfterSpecificBookmark();
        }

        /// <summary>
        /// How to insert a table in DOCX document after a specific bookmark.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-insert-table-in-docx-after-specific-bookmark-net-csharp-vb.php
        /// </remarks>		
        public static void InsertTableAfterSpecificBookmark()
        {
            // A one of our customers sent us a request to help him with such example. 
            // He has a .docx document with several bookmarks, he wants to insert 
            // a new table after the specific bookmark. Let's see how to make it.

            string inpFile = @"..\..\..\bookmarks.docx";
            string outFile = @"Result.docx";

            // 1. Load a document with bookmarks.
            DocumentCore dc = DocumentCore.Load(inpFile);

            // 2. Find a specific bookmark by name.
            // Our DOCX document contains 2 bookmarks: table1 and table2.
            Bookmark bookmark = dc.Bookmarks["table2"];

            // 3. Insert a table after the bookmark "table2".
            if (bookmark != null)
            {
                // Create a new simple table 2 x 3.
                Table table = CreateTable(dc, 2, 3, 100f);

                // Insert the table after the bookmark.
                bookmark.End.Content.End.Insert(table.Content);
            }

            // Let's save our document into DOCX format.
            dc.Save(outFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }

        public static Table CreateTable(DocumentCore dc, int rows, int columns, float widthMm)
        {
            // Let's create a plain table: 2x3, 100 mm of width.
            Table table = new Table(dc);
            double width = LengthUnitConverter.Convert(widthMm, LengthUnit.Millimeter, LengthUnit.Point);
            table.TableFormat.PreferredWidth = new TableWidth(width, TableWidthUnit.Point);
            table.TableFormat.Alignment = HorizontalAlignment.Center;

            int counter = 0;

            // Add rows.
            for (int r = 0; r < rows; r++)
            {
                TableRow row = new TableRow(dc);

                // Add columns.
                for (int c = 0; c < columns; c++)
                {
                    TableCell cell = new TableCell(dc);

                    // Set cell formatting and width.
                    cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dotted, Color.Black, 1.0);

                    // Set the same width for each column.
                    cell.CellFormat.PreferredWidth = new TableWidth(width / columns, TableWidthUnit.Point);

                    if (counter % 2 == 1)
                        cell.CellFormat.BackgroundColor = new Color("#358CCB");

                    row.Cells.Add(cell);

                    // Let's add a paragraph with text into the each column.
                    Paragraph p = new Paragraph(dc);
                    p.ParagraphFormat.Alignment = HorizontalAlignment.Center;
                    p.ParagraphFormat.SpaceBefore = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point);
                    p.ParagraphFormat.SpaceAfter = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point);

                    p.Content.Start.Insert(String.Format("{0}", (char)(counter + 'A')), new CharacterFormat() { FontName = "Arial", FontColor = new Color("#3399FF"), Size = 12.0 });
                    cell.Blocks.Add(p);
                    counter++;
                }
                table.Rows.Add(row);
            }
            return table;
        }
    }
}