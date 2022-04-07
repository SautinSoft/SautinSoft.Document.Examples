Imports System.IO
Imports SautinSoft.Document
Imports System.Linq
Imports System.Text.RegularExpressions

Module Sample
    Sub Main()
        FindAndReplace()
    End Sub

    ''' <summary>
    ''' Find and replace a text using ContentRange.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/find-replace-content-net-csharp-vb.php
    ''' </remarks>
    Sub FindAndReplace()
        ' Path to a loadable document.
        Dim loadPath As String = "..\..\..\critique.docx"

        ' Load a document intoDocumentCore.
        Dim dc As DocumentCore = DocumentCore.Load(loadPath)

        Dim regex As New Regex("bean", RegexOptions.IgnoreCase)

        'Find "Bean" and Replace everywhere on "Joker :-)"
        ' Please note, Reverse() makes sure that action replace not affects to Find().
        For Each item As ContentRange In dc.Content.Find(regex).Reverse()
            item.Replace("Joker")
        Next item

        ' Save our document into PDF format.
        Dim savePath As String = "Replaced.pdf"
        dc.Save(savePath, New PdfSaveOptions())

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(loadPath) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(savePath) With {.UseShellExecute = True})
    End Sub
End Module