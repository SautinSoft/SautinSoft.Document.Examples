using System;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            AddSimpleTable();
        }
        
		/// <summary>
        /// How to create a plain table in a document. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/add-simple-table.php
        /// </remarks>
        public static void AddSimpleTable()
        {
            string documentPath = @"SimpleTable.pdf";

            // Let's create a new document.
            DocumentCore dc = new DocumentCore();

            // Add a new section.
            Section s = new Section(dc);
            dc.Sections.Add(s);

            // Let's create a plain table: 2x3, 100 mm of width.
            Table table = new Table(dc);
            double width = LengthUnitConverter.Convert(100, LengthUnit.Millimeter, LengthUnit.Point);
            table.TableFormat.PreferredWidth = new TableWidth(width, TableWidthUnit.Point);
            table.TableFormat.Alignment = HorizontalAlignment.Center;

            int counter = 0;

            // Add rows.
            int rows = 2;
            int columns = 3;
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

            // Add the table into the section.
            s.Blocks.Add(table);

            // Save our document into PDF format.
            dc.Save(documentPath, new PdfSaveOptions() { Compliance = PdfCompliance.PDF_A1a });

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}