Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document

Module Sample
    Sub Main()
        FindAndReplaceInParagraphs()
    End Sub

    ''' <summary>
    ''' Find and replace a specific text in all paragraphs in PDF document.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/replace-text-paragraphs-in-pdf-document-net-csharp-vb.php
    ''' </remarks>
    Private Sub FindAndReplaceInParagraphs()
        Dim filePath As String = "..\..\..\example.pdf"
        Dim fileResult As String = "Result.pdf"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)
        For Each par As Paragraph In dc.GetChildElements(True, ElementType.Paragraph)
            For Each item As ContentRange In par.Content.Find("old text").Reverse()
                item.Replace("new text")
            Next item
        Next par
        dc.Save(fileResult)
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(fileResult) With {.UseShellExecute = True})
    End Sub
End Module