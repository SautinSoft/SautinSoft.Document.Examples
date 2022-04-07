Imports System
Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Tables

Module Sample
    Sub Main()
        AddSimpleTable()
    End Sub

    ''' <summary>
    ''' How to create a plain table in a document. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/add-simple-table.php
    ''' </remarks>
    Sub AddSimpleTable()
        Dim documentPath As String = "SimpleTable.pdf"

        ' Let's create a new document.
        Dim dc As New DocumentCore()

        ' Add a new section.
        Dim s As New Section(dc)
        dc.Sections.Add(s)

        ' Let's create a plain table: 2x3, 100 mm of width.
        Dim table As New Table(dc)
        Dim width As Double = LengthUnitConverter.Convert(100, LengthUnit.Millimeter, LengthUnit.Point)
        table.TableFormat.PreferredWidth = New TableWidth(width, TableWidthUnit.Point)
        table.TableFormat.Alignment = HorizontalAlignment.Center

        Dim counter As Integer = 0

        ' Add rows.
        Dim rows As Integer = 2
        Dim columns As Integer = 3
        For r As Integer = 0 To rows - 1
            Dim row As New TableRow(dc)

            ' Add columns.
            For c As Integer = 0 To columns - 1
                Dim cell As New TableCell(dc)

                ' Set cell formatting and width.
                cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dotted, Color.Black, 1.0)

                ' Set the same width for each column.
                cell.CellFormat.PreferredWidth = New TableWidth(width / columns, TableWidthUnit.Point)

                If counter Mod 2 = 1 Then
                    cell.CellFormat.BackgroundColor = New Color("#358CCB")
                End If

                row.Cells.Add(cell)

                ' Let's add a paragraph with text into the each column.
                Dim p As New Paragraph(dc)
                p.ParagraphFormat.Alignment = HorizontalAlignment.Center
                p.ParagraphFormat.SpaceBefore = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point)
                p.ParagraphFormat.SpaceAfter = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point)

                p.Content.Start.Insert(String.Format("{0}", ChrW(counter + AscW("A"c))), New CharacterFormat() With {
                    .FontName = "Arial",
                    .FontColor = New Color("#3399FF"),
                    .Size = 12.0
                })
                cell.Blocks.Add(p)
                counter += 1
            Next c
            table.Rows.Add(row)
        Next r

        ' Add the table into the section.
        s.Blocks.Add(table)

        ' Save our document into PDF format.
        dc.Save(documentPath, New PdfSaveOptions() With {.Compliance = PdfCompliance.PDF_A1a})

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub
End Module