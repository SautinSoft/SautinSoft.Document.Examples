Imports System.IO
Imports System.Linq
Imports SautinSoft.Document
Imports SautinSoft.Document.Tables

Namespace Sample
    Friend Class Sample
        Shared Sub Main(ByVal args() As String)
            MergeRowsAndCellsInTable()
        End Sub

        ''' <summary>
        ''' Create a new table with rows and cells merged by vertical (rowspan) and horizontal (colspan).
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/table-with-merged-rows-and-cells.php
        ''' </remarks>
        Public Shared Sub MergeRowsAndCellsInTable()

            Dim dc As New DocumentCore()

            Dim table As Table = New Table(dc,
                 New TableRow(dc,
                         New TableCell(dc, New Paragraph(dc, "Cell 1-1")),
                         New TableCell(dc, New Paragraph(dc, "Cell 1-2")),
                         New TableCell(dc, New Paragraph(dc, "Cell 1-3")),
                         New TableCell(dc, New Paragraph(dc, "Cell 1-4"))),
                 New TableRow(dc,
                         New TableCell(dc, New Paragraph(dc, "Cell 2-1 -> 3-2")) With
                {
                .RowSpan = 2,
                .ColumnSpan = 2
                },
                         New TableCell(dc, New Paragraph(dc, "Cell 2-3 -> 2-4")) With
                {.ColumnSpan = 2}),
                New TableRow(dc,
                         New TableCell(dc) With
                {.ColumnSpan = 2},
                         New TableCell(dc, New Paragraph(dc, "Cell 3-3")),
                         New TableCell(dc, New Paragraph(dc, "Cell 3-4"))),
                New TableRow(dc, New TableCell(dc, New Paragraph(dc, "Cell 4-1"))),
                New TableRow(dc, New TableCell(dc, New Paragraph(dc, "Cell 5-1"))))

            table.TableFormat.DefaultCellPadding = New Padding(10, LengthUnit.Pixel)

            ' Set the table width to 10 cm and convert it to points.
            Dim tableWidthInPoints As Double = LengthUnitConverter.Convert(10, LengthUnit.Centimeter, LengthUnit.Point)
            table.TableFormat.PreferredWidth = New TableWidth(tableWidthInPoints, TableWidthUnit.Point)
            For r As Integer = 0 To table.Rows.Count - 1
                Dim c As Integer = 0
                Do While c < table.Rows(r).Cells.Count
                    Dim cell As TableCell = table.Rows(r).Cells(c)
                    cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dashed, Color.Black, 1)
                    cell.CellFormat.BackgroundColor = New Color("#FFCC00")
                    c += 1
                Loop
            Next r

            dc.Sections.Add(New Section(dc, table))

            dc.Save("MergedTableCells.docx")
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo("MergedTableCells.docx") With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
