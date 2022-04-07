Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        SplitDocumentByPages()
    End Sub

    ''' <summary>
    ''' Loads a document and split it by separate pages. Saves the each page into PDF format.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/split-docx-document-by-pages-in-pdf-format-net-csharp-vb.php
    ''' </remarks>
    Sub SplitDocumentByPages()
        Dim filePath As String = "..\..\..\example.docx"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)
        Dim folderPath As String = Path.GetFullPath("Result-files")
        Dim dp As DocumentPaginator = dc.GetPaginator()
        For i As Integer = 0 To dp.Pages.Count - 1
            Dim page As DocumentPage = dp.Pages(i)
            Directory.CreateDirectory(folderPath)

            ' Save the each page to PDF format.
            page.Save(folderPath & "\Page - " & i.ToString() & ".pdf", SaveOptions.PdfDefault)
        Next i
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(folderPath) With {.UseShellExecute = True})
    End Sub
End Module