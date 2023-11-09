Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        SplitDocumentByPages()
    End Sub

     Sub SplitDocumentByPages()
        Dim filePath As String = "..\..\..\example.docx"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)
        Dim folderPath As String = Path.GetFullPath("Result-files")
        Dim dp As DocumentPaginator = dc.GetPaginator()
        For i As Integer = 0 To dp.Pages.Count - 1
            Dim page As DocumentPage = dp.Pages(i)
            Directory.CreateDirectory(folderPath)

            ' Save the each page to DOCX format.
            page.Save(folderPath & "\Page - " & i.ToString() & ".docx", SaveOptions.PdfDefault)
        Next i
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(folderPath) With {.UseShellExecute = True})
    End Sub
End Module