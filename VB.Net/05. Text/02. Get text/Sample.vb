Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        GetText()
    End Sub

    ''' <summary>
    ''' Get all Text (Run objects) from DOCX document and show it on Console.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/get-text-from-docx-document-net-csharp-vb.php
    ''' </remarks>
    Sub GetText()
        Dim filePath As String = "..\..\..\example.docx"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)

        ' Get all Run elements from document.
        For Each run As Run In dc.GetChildElements(True, ElementType.Run)
            Console.WriteLine(run.Text)
        Next run
        Console.ReadKey()
    End Sub
End Module