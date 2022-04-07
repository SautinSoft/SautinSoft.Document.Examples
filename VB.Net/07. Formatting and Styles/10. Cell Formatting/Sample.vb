Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports SautinSoft.Document.Tables
Imports System.Linq

Namespace Example
    Friend Class Program
        Shared Sub Main(ByVal args() As String)
            TableCellFormatting()
        End Sub
        ''' <summary>
        ''' How apply formatting for table cells.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/tablecell-format.php
        ''' </remarks>
        Private Shared Sub TableCellFormatting()
            Dim filePath As String = "TableCellFormatting.docx"

            ' Let's a new create document.
            Dim dc As New DocumentCore()

            ' Add new table.
            Dim table As New Table(dc)
            Dim width As Double = LengthUnitConverter.Convert(100, LengthUnit.Millimeter, LengthUnit.Point)
            table.TableFormat.PreferredWidth = New TableWidth(width, TableWidthUnit.Point)
            table.TableFormat.Borders.SetBorders(MultipleBorderTypes.Inside Or MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 1)
            dc.Sections.Add(New Section(dc, table))

            Dim row As New TableRow(dc)
            row.RowFormat.Height = New TableRowHeight(50, HeightRule.Exact)
            table.Rows.Add(row)

            ' Create a cells with formatting: borders, background color, vertical alignment and text direction.
            Dim cell As New TableCell(dc, New Paragraph(dc, "Cell 1"))
            cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Blue, 1)
            cell.CellFormat.BackgroundColor = Color.Yellow
            row.Cells.Add(cell)

            row.Cells.Add(New TableCell(dc, New Paragraph(dc, "Cell 2")) With {
                .CellFormat = New TableCellFormat() With {.VerticalAlignment = VerticalAlignment.Center}
            })

            row.Cells.Add(New TableCell(dc, New Paragraph(dc, "Cell 3")) With {
                .CellFormat = New TableCellFormat() With {.TextDirection = TextDirection.BottomToTop}
            })

            ' Save our document.
            dc.Save(filePath)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace