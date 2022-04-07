Imports System
Imports System.Linq
Imports SautinSoft.Document
Module Sample
    Sub Main()
        ShowLists()
    End Sub

    ''' <summary>
    ''' Find all paragraphs in a document marked as list (ordered or unordered).
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/elementcollection-linq.php
    ''' </remarks>
    Sub ShowLists()
        Dim filePath As String = "..\..\..\example.docx"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)

        ' Select all paragraphs marked as list using LINQ.
        Dim selectedPars = From p In dc.GetChildElements(True, ElementType.Paragraph)
                           Where (TryCast(p, Paragraph)).ListFormat.IsList
                           Select p

        For Each p As Paragraph In selectedPars
            Console.WriteLine(p.Content.ToString().TrimEnd())
        Next p

        Console.ReadKey()
    End Sub
End Module