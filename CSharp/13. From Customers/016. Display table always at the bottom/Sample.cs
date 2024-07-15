using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using SautinSoft.Document.Tables;
using System;

// Get your free 100-day key here:   
// https://sautinsoft.com/start-for-free/
DisplayTable();

/// <summary>
/// How to display a table always at the bottom of the page.
/// </summary>
/// <remarks>
/// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-display-table-always-at-the-bottom-of-the-page-in-pdf-net-csharp-vb.php
/// </remarks>
static void DisplayTable()
{
    // Is there a way to display a table always at the bottom of the page?
    // Yes, place the table in the document footer.

    // Create 5-pages document and place an unique table
    // at the bottom of each page.
    var dc = new DocumentCore();

    string resultPath = "Result.pdf";
    string[] pagesText = { "March", "April", "May", "June", "July" };

    for (int page = 0; page < pagesText.Length; page++)
    {
        var s = new Section(dc);
        dc.Sections.Add(s);
        // Write some text content
        var p = new Paragraph(dc);
        p.ParagraphFormat.Alignment = HorizontalAlignment.Center;
        var cf = new CharacterFormat() { Size = 100f };
        p.Inlines.Add(new Run(dc, pagesText[page], cf));
        s.Blocks.Add(p);

        // Place the table in the document footer.                
        AddTableToFooter(dc, s);
    }

    dc.Save(resultPath);
    // Open the result for demonstration purposes.
    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
}
static void AddTableToFooter(DocumentCore dc, Section s)
{
    Random rand = new Random();
    Paragraph p;

    string[] tableText = { "Item", "Quantity", "Price" };
    Table t = new Table(dc);
    t.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);
    t.TableFormat.Alignment = HorizontalAlignment.Center;

    // Table header
    var rowHdr = new TableRow(dc);
    foreach (var cellText in tableText)
    {
        var cellHdr = new TableCell(dc);
        cellHdr.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 0.5f);
        cellHdr.CellFormat.BackgroundColor = new Color(rand.Next(Int32.MaxValue));
        cellHdr.CellFormat.PreferredWidth = new TableWidth(100.0f / 3.0f, TableWidthUnit.Percentage);
        p = new Paragraph(dc, cellText);
        p.ParagraphFormat.Alignment = HorizontalAlignment.Center;
        cellHdr.Blocks.Add(p);
        rowHdr.Cells.Add(cellHdr);
    }
    t.Rows.Add(rowHdr);
    // Table body
    int rowCount = rand.Next(1, 20);
    for (int r = 0; r < rowCount; r++)
    {
        var row = new TableRow(dc);
        foreach (var cellText in tableText)
        {
            var cell = new TableCell(dc);
            cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 0.5f);
            cell.CellFormat.BackgroundColor = Color.White;
            cell.CellFormat.PreferredWidth = new TableWidth(100.0f / 3.0f, TableWidthUnit.Percentage);
            p = new Paragraph(dc, $"{rand.Next(100)}");
            p.ParagraphFormat.Alignment = HorizontalAlignment.Center;
            cell.Blocks.Add(p);
            row.Cells.Add(cell);
        }
        t.Rows.Add(row);
    }
    // Move table to page footer
    HeaderFooter footer = new HeaderFooter(dc, HeaderFooterType.FooterDefault);
    s.HeadersFooters.Add(footer);
    footer.Blocks.Add(t);
}