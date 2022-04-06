using System;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            AddNestedTable();
        }

		/// <summary>
        /// How to create a nested table in a document. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/add-nested-table.php
        /// </remarks>
        public static void AddNestedTable()
        {
            string documentPath = @"NestedTable.docx";

            // Let's create a new document.
            DocumentCore dc = new DocumentCore();

            // Add a new section.
            Section s = new Section(dc);
            dc.Sections.Add(s);

            // Let's create a table1: 1x2, with 150 mm width.
            Table table1 = new Table(dc);
            double twidth = LengthUnitConverter.Convert(150, LengthUnit.Millimeter, LengthUnit.Point);
            table1.TableFormat.PreferredWidth = new TableWidth(twidth, TableWidthUnit.Point);

            // Add 1 rows.
            for (int r = 0; r < 1; r++)
            {
                TableRow row = new TableRow(dc);

                // Add 2 columns.
                for (int c = 0; c < 2; c++)
                {
                    TableCell cell = new TableCell(dc);

                    // Set cell formatting and width.
                    cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 1.0);
                    cell.CellFormat.PreferredWidth = new TableWidth(twidth / 2, TableWidthUnit.Point);

                    double padding = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point);
                    cell.CellFormat.Padding = new Padding(padding);

                    row.Cells.Add(cell);
                }
                table1.Rows.Add(row);
            }

            // Add this table to the current section.
            s.Blocks.Add(table1);

            // Create nested table2 3x3.
            Table table2 = new Table(dc);
            twidth = LengthUnitConverter.Convert(75, LengthUnit.Millimeter, LengthUnit.Point);
            table2.TableFormat.PreferredWidth = new TableWidth(twidth, TableWidthUnit.Point);
            table2.TableFormat.Alignment = HorizontalAlignment.Center;

            for (int r = 0; r < 3; r++)
            {
                TableRow row = new TableRow(dc);

                // Add 2 columns
                for (int c = 0; c < 3; c++)
                {
                    TableCell cell = new TableCell(dc);

                    // Set cell formatting and width.
                    cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 1.0);
                    if (c % 2 == 0)
                        cell.CellFormat.BackgroundColor = Color.Orange;
                    else
                        cell.CellFormat.BackgroundColor = new Color("#358CCB");

                    cell.CellFormat.PreferredWidth = new TableWidth(twidth / 2, TableWidthUnit.Point);

                    row.Cells.Add(cell);

                    // Let's add some text into each column.
                    Paragraph p = new Paragraph(dc);
                    p.ParagraphFormat.Alignment = HorizontalAlignment.Center;
                    p.ParagraphFormat.SpaceBefore = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point);
                    p.ParagraphFormat.SpaceAfter = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point); ;

                    p.Content.Start.Insert(String.Format("({0},{1})", r + 1, c + 1, new CharacterFormat() { FontName = "Arial", Size = 12.0 }));
                    cell.Blocks.Add(p);
                }
                table2.Rows.Add(row);
            }

            // Insert table2 inside 2nd columns of table 1.
            table1.Rows[0].Cells[1].Blocks.Add(table2);

            // Insert some text inside 1st column of table 1.
            Paragraph p2 = new Paragraph(dc);
            p2.ParagraphFormat.Alignment = HorizontalAlignment.Center;
            p2.Content.Start.Insert("This is a 1st column of table 1");
            table1.Rows[0].Cells[0].Blocks.Add(p2);

            // Save our document into DOCX format.
            dc.Save(documentPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}