Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document
Imports SautinSoft.Document.Tables

Module Sample
    Sub Main()
        ModifyTable()
    End Sub

    ''' <summary>
    ''' How to modify an existing table in a document.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/modify-table.php
    ''' </remarks>
    Sub ModifyTable()
        Dim sourcePath As String = "..\..\..\table.docx"
        Dim destPath As String = "Table modified.docx"

        ' Load a document with a table.
        Dim dc As DocumentCore = DocumentCore.Load(sourcePath)

        ' Find a first table in the document.
        Dim table As Table = CType(dc.GetChildElements(True, ElementType.Table).First(), Table)

        ' Set dashed borders and yellow background for all cells.
        For r As Integer = 0 To table.Rows.Count - 1
            Dim c As Integer = 0
            Do While c < table.Rows(r).Cells.Count
                Dim cell As TableCell = table.Rows(r).Cells(c)
                cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dashed, Color.Black, 1)
                cell.CellFormat.BackgroundColor = New Color("#FFCC00")
                c += 1
            Loop
        Next r

        ' Save the document as DOCX.
        dc.Save(destPath, New DocxSaveOptions())

        ' Show the source and the dest documents.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(sourcePath) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(destPath) With {.UseShellExecute = True})
    End Sub
End Module