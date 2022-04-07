Imports System
Imports SautinSoft.Document
Imports SautinSoft.Document.Tables

Module Sample
    Sub Main()
        TableFormat()
    End Sub

    ''' <summary>
    ''' How to create a table and apply formatting in a document. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/table-format.php
    ''' </remarks>
    Sub TableFormat()
        Dim documentPath As String = "TableFormat.docx"

        ' Let's create a new document.
        Dim dc As New DocumentCore()

        ' Add a new section.
        Dim s As New Section(dc)
        dc.Sections.Add(s)


        Dim table As New Table(dc)
        table.TableFormat.AutomaticallyResizeToFitContents = False
        table.TableFormat.Alignment = HorizontalAlignment.Center
        table.TableFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dotted, Color.DarkBlue, 2.0)

        table.Columns.Add(New TableColumn() With {.PreferredWidth = 50})
        table.Columns.Add(New TableColumn() With {.PreferredWidth = 80})
        table.Columns.Add(New TableColumn() With {.PreferredWidth = 100})
        table.Columns.Add(New TableColumn() With {.PreferredWidth = 120})
        table.Columns.Add(New TableColumn() With {.PreferredWidth = 80})

        Dim row As New TableRow(dc)
        row.RowFormat.Height = New TableRowHeight(100, HeightRule.AtLeast)
        table.Rows.Add(row)

        Dim cell1 As New TableCell(dc, New Paragraph(dc, "PDF is Portable Document Format"))
        cell1.CellFormat.TextDirection = TextDirection.TopToBottom
        cell1.CellFormat.BackgroundColor = Color.Yellow
        cell1.CellFormat.Padding = New Padding(5)
        cell1.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dotted, Color.Yellow, 1.0)
        row.Cells.Add(cell1)

        Dim cell2 As New TableCell(dc, New Paragraph(dc, "DOCX is Office Open XML"))
        cell2.CellFormat.VerticalAlignment = VerticalAlignment.Center
        cell2.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dotted, Color.Black, 1.0)
        row.Cells.Add(cell2)

        row.Cells.Add(New TableCell(dc, New Paragraph(dc, "HTML is Hypertext Markup Language")) With {
            .CellFormat = New TableCellFormat() With {.BackgroundColor = Color.Pink}
        })

        row.Cells.Add(New TableCell(dc, New Paragraph(dc, "Images: jpeg, png, bmp, tiff") With
        {
            .ParagraphFormat = New ParagraphFormat() With
            {.Alignment = HorizontalAlignment.Center}
        }) With {
            .CellFormat = New TableCellFormat() With {
                .VerticalAlignment = VerticalAlignment.Center,
                .BackgroundColor = Color.Purple
            }
        })

        Dim cell5 As New TableCell(dc, New Paragraph(dc, "RTF is Rich Text Format"))
        cell5.CellFormat.TextDirection = TextDirection.BottomToTop
        cell5.CellFormat.Padding = New Padding(5)
        cell5.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dashed, Color.Red, 1.0)
        row.Cells.Add(cell5)

        ' Add the table into the section.
        s.Blocks.Add(table)

        ' Save our document into DOCX format.
        dc.Save(documentPath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub

End Module