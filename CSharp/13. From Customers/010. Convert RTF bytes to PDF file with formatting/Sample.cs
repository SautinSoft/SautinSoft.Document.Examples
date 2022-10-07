using System;
using System.IO;
using SautinSoft.Document;
using SautinSoft.Document.Tables;
using System.Linq;
using System.Text.RegularExpressions;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            ConvertRtfBytesToPdfFile();
        }
        public static void ConvertRtfBytesToPdfFile()
        {
            // Get document bytes.
            byte[] fileBytes = File.ReadAllBytes(@"..\..\..\example.rtf");
            
            string PdfPath = @"result.pdf";
            
            DocumentCore dc = null;
            Regex regex = new Regex(@"formatting", RegexOptions.IgnoreCase);
            
            // Create a MemoryStream.
            using (MemoryStream ms = new MemoryStream(fileBytes))
            {
                // Load a document from the MemoryStream.
                // Specifying RtfLoadOptions we explicitly set that a loadable document is RTF.
                dc = DocumentCore.Load(ms, new RtfLoadOptions());
            }

            // Add a new section in the document.
            Section s = new Section(dc);
            dc.Sections.Add(s);

            // Create a new table with two rows and three columns inside.
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
                    cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dashed, Color.Brown, 2.0);

                    // Set the same width for each column.
                    cell.CellFormat.PreferredWidth = new TableWidth(width / columns, TableWidthUnit.Point);

                    if (counter % 2 == 1)
                        cell.CellFormat.BackgroundColor = new Color("#FF0000");

                    row.Cells.Add(cell);

                    // Let's add a paragraph with text into the each column.
                    Paragraph pa = new Paragraph(dc);
                    pa.ParagraphFormat.Alignment = HorizontalAlignment.Center;
                    pa.ParagraphFormat.SpaceBefore = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point);
                    pa.ParagraphFormat.SpaceAfter = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point);

                    pa.Content.Start.Insert(String.Format("{0}", (char)(counter + 'A')), new CharacterFormat() { FontName = "Arial", FontColor = new Color("#000000"), Size = 12.0 });
                    cell.Blocks.Add(pa);
                    counter++;
                }
                table.Rows.Add(row);

            }
                // Create a new header with formatted text.
                HeaderFooter header = new HeaderFooter(dc, HeaderFooterType.HeaderDefault);
                header.Content.Start.Insert(table.Content);
                foreach (Section s1 in dc.Sections)
                {
                    s1.HeadersFooters.Add(header.Clone(true));
                }

                // Add the header into HeadersFooters collection of the 1st section.
                //s1.HeadersFooters.Add(header);

                // Create a new footer with formatted text.
                HeaderFooter footer = new HeaderFooter(dc, HeaderFooterType.FooterDefault);
                footer.Content.Start.Insert(table.Content);
                foreach (Section s1 in dc.Sections)
                {
                    s1.HeadersFooters.Add(footer.Clone(true));
                }

                // Add the footer into HeadersFooters collection of the 1st section.
                //s1.HeadersFooters.Add(footer);

                foreach (ContentRange item in dc.Content.Find(regex).Reverse())
                {
                // Replace all text "formatting" on "FORMATTING!!!".
                item.Replace("FORMATTING!!!", new CharacterFormat() { BackgroundColor = Color.Yellow, FontName = "Arial", Size = 16.0 });
                }

                // Save our result as a PDF file.
            dc.Save(PdfPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(PdfPath) { UseShellExecute = true });
        }
    }
}  

