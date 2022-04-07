Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document

Module Sample
    Sub Main()
        DeletePageBreak()
    End Sub

    ''' <summary>
    ''' Working with special characters in a document. How delete all page breaks in DOCX.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/special-character-text-in-docx-document-net-csharp-vb.php
    ''' </remarks>
    Sub DeletePageBreak()
        Dim filePath As String = "..\..\..\example.docx"
        Dim fileResult As String = "Result.docx"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)
        For Each sc As SpecialCharacter In dc.GetChildElements(True, ElementType.SpecialCharacter).Reverse()
            If sc.CharacterType = SpecialCharacterType.PageBreak Then
                sc.Parent.Content.Delete()
            End If
        Next sc
        dc.Save(fileResult)
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(fileResult) With {.UseShellExecute = True})
    End Sub
End Module