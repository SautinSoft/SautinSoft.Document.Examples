Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        GetParagraphs()
    End Sub

    ''' <summary>
    ''' Loads an existing DOCX document and renders all paragraphs to Console.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/get-paragraphs-from-docx-document-net-csharp-vb.php
    ''' </remarks>
    Sub GetParagraphs()
        Dim filePath As String = "..\..\..\example.docx"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)
        For Each par As Paragraph In dc.GetChildElements(True, ElementType.Paragraph)
            Console.WriteLine(par.Content.ToString())
        Next par
        Console.ReadKey()
    End Sub
End Module