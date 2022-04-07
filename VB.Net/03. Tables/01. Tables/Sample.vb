Imports SautinSoft.Document
Imports SautinSoft.Document.Tables

Namespace Example
    Friend Class Program
        Private Shared dc As DocumentCore

        Shared Sub Main(ByVal args() As String)
            CreateTable()
        End Sub

        ''' <summary>
        ''' Creates a new cell within a table and fills the it in staggered order.
        ''' </summary>
        ''' <param name="rowIndex">Row index.</param>
        ''' <param name="colIndex">Column index.</param>
        ''' <returns>New table cell.</returns>
        Public Shared Function NewCell(ByVal rowIndex As Integer, ByVal colIndex As Integer) As TableCell
            Dim cell As New TableCell(dc)

            cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 1)

            If colIndex Mod 2 = 1 AndAlso rowIndex Mod 2 = 0 OrElse colIndex Mod 2 = 0 AndAlso rowIndex Mod 2 = 1 Then
                cell.CellFormat.BackgroundColor = Color.Black
            End If

            Dim run As New Run(dc, String.Format("Row - {0}; Col - {1}", rowIndex, colIndex))
            run.CharacterFormat.FontColor = Color.Auto

            cell.Blocks.Content.Replace(run.Content)

            Return cell
        End Function

        ''' <summary>
        ''' Creates a new document with a table.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/tables.php
        ''' </remarks>
        Public Shared Sub CreateTable()
            dc = New DocumentCore()

            Dim filePath As String = "Result-file.docx"

            Dim table As New Table(dc, 5, 5, AddressOf NewCell)

            ' Place the 'Table' at the start of the 'Document'.
            ' By the way, we didn't create a 'Section' in our document.
            ' As we're using 'Content' property, a 'Section' will be created automatically if necessary.
            dc.Content.Start.Insert(table.Content)
            dc.Save(filePath)

            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
