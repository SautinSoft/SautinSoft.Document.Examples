Imports System
Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Tables

Module Sample
    Sub Main()
        AddNestedTable()
    End Sub

    ''' <summary>
    ''' How to create a nested table in a document. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/add-nested-table.php
    ''' </remarks>
    Sub AddNestedTable()
        Dim documentPath As String = "NestedTable.docx"

        ' Let's create a new document.
        Dim dc As New DocumentCore()

        ' Add a new section.
        Dim s As New Section(dc)
        dc.Sections.Add(s)

        ' Let's create a table1: 1x2, with 150 mm width.
        Dim table1 As New Table(dc)
        Dim twidth As Double = LengthUnitConverter.Convert(150, LengthUnit.Millimeter, LengthUnit.Point)
        table1.TableFormat.PreferredWidth = New TableWidth(twidth, TableWidthUnit.Point)

        ' Add 1 rows.
        For r As Integer = 0 To 0
            Dim row As New TableRow(dc)

            ' Add 2 columns.
            For c As Integer = 0 To 1
                Dim cell As New TableCell(dc)

                ' Set cell formatting and width.
                cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 1.0)
                cell.CellFormat.PreferredWidth = New TableWidth(twidth / 2, TableWidthUnit.Point)

                Dim padding As Double = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point)
                cell.CellFormat.Padding = New Padding(padding)

                row.Cells.Add(cell)
            Next c
            table1.Rows.Add(row)
        Next r

        ' Add this table to the current section.
        s.Blocks.Add(table1)

        ' Create nested table2 3x3.
        Dim table2 As New Table(dc)
        twidth = LengthUnitConverter.Convert(75, LengthUnit.Millimeter, LengthUnit.Point)
        table2.TableFormat.PreferredWidth = New TableWidth(twidth, TableWidthUnit.Point)
        table2.TableFormat.Alignment = HorizontalAlignment.Center

        For r As Integer = 0 To 2
            Dim row As New TableRow(dc)

            ' Add 2 columns
            For c As Integer = 0 To 2
                Dim cell As New TableCell(dc)

                ' Set cell formatting and width.
                cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 1.0)
                If c Mod 2 = 0 Then
                    cell.CellFormat.BackgroundColor = Color.Orange
                Else
                    cell.CellFormat.BackgroundColor = New Color("#358CCB")
                End If

                cell.CellFormat.PreferredWidth = New TableWidth(twidth / 2, TableWidthUnit.Point)

                row.Cells.Add(cell)

                ' Let's add some text into each column.
                Dim p As New Paragraph(dc)
                p.ParagraphFormat.Alignment = HorizontalAlignment.Center
                p.ParagraphFormat.SpaceBefore = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point)
                p.ParagraphFormat.SpaceAfter = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point)


                p.Content.Start.Insert(String.Format("({0},{1})", r + 1, c + 1, New CharacterFormat() With {
                    .FontName = "Arial",
                    .Size = 12.0
                }))
                cell.Blocks.Add(p)
            Next c
            table2.Rows.Add(row)
        Next r

        ' Insert table2 inside 2nd columns of table 1.
        table1.Rows(0).Cells(1).Blocks.Add(table2)

        ' Insert some text inside 1st column of table 1.
        Dim p2 As New Paragraph(dc)
        p2.ParagraphFormat.Alignment = HorizontalAlignment.Center
        p2.Content.Start.Insert("This is a 1st column of table 1")
        table1.Rows(0).Cells(0).Blocks.Add(p2)

        ' Save our document into DOCX format.
        dc.Save(documentPath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub
End Module