Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document

Module Sample
    Sub Main()
        FindText()
    End Sub

    ''' <summary>
    ''' Find a specific text in DOCX document.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/find-text-in-docx-document-net-csharp-vb.php
    ''' </remarks>
    Sub FindText()
        Dim filePath As String = "..\..\..\example.docx"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)
        Dim searchText As String = "document"
        Dim count As Integer = dc.Content.Find(searchText).Count()

        Console.WriteLine("The text: """ & searchText & """ - was found " & count.ToString() & " time(s).")
        Console.ReadKey()
    End Sub
End Module