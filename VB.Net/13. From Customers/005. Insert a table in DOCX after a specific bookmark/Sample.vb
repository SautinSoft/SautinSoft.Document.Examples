Imports SautinSoft.Document
Imports SautinSoft.Document.Tables
Imports System

Namespace Sample
    Friend Class Sample
        Shared Sub Main(ByVal args() As String)
            InsertTableAfterSpecificBookmark()
        End Sub

        ''' <summary>
        ''' How to insert a table in DOCX document after a specific bookmark.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-insert-table-in-docx-after-specific-bookmark-net-csharp-vb.php
        ''' </remarks>		
        Public Shared Sub InsertTableAfterSpecificBookmark()
            ' A one of our customers sent us a request to help him with such example. 
            ' He has a .docx document with several bookmarks, he wants to insert 
            ' a new table after the specific bookmark. Let's see how to make it.

            Dim inpFile As String = "..\..\..\bookmarks.docx"
            Dim outFile As String = "Result.docx"

            ' 1. Load a document with bookmarks.
            Dim dc As DocumentCore = DocumentCore.Load(inpFile)

            ' 2. Find a specific bookmark by name.
            ' Our DOCX document contains 2 bookmarks: table1 and table2.
            Dim bookmark As Bookmark = dc.Bookmarks("table2")

            ' 3. Insert a table after the bookmark "table2".
            If bookmark IsNot Nothing Then
                ' Create a new simple table 2 x 3.
                Dim table As Table = CreateTable(dc, 2, 3, 100.0F)

                ' Insert the table after the bookmark.
                bookmark.End.Content.End.Insert(table.Content)
            End If

            ' Let's save our document into DOCX format.
            dc.Save(outFile)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
        End Sub

        Public Shared Function CreateTable(ByVal dc As DocumentCore, ByVal rows As Integer, ByVal columns As Integer, ByVal widthMm As Single) As Table
            ' Let's create a plain table: 2x3, 100 mm of width.
            Dim table As New Table(dc)
            Dim width As Double = LengthUnitConverter.Convert(widthMm, LengthUnit.Millimeter, LengthUnit.Point)
            table.TableFormat.PreferredWidth = New TableWidth(width, TableWidthUnit.Point)
            table.TableFormat.Alignment = HorizontalAlignment.Center

            Dim counter As Integer = 0

            ' Add rows.
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
            Return table
        End Function
    End Class
End Namespace
