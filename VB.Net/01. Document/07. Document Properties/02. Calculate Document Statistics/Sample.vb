Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        CalculateStatistics()
    End Sub

    ''' <summary>
    ''' Calculates the number of words, pages and characters in a document.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/counting-words-paragraphs-net-csharp-vb.php
    ''' </remarks>
    Sub CalculateStatistics()
        ' Load a DOCX file.
        Dim filePath As String = "..\..\..\words.docx"

        Dim dc As DocumentCore = DocumentCore.Load(filePath)

        ' Update and count the number of words and pages in the file.
        dc.CalculateStats()

        ' Show statistics.
        Console.WriteLine("Pages: {0}", dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.Pages))
        Console.WriteLine("Paragraphs: {0}", dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.Paragraphs))
        Console.WriteLine("Words: {0}", dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.Words))
        Console.WriteLine("Characters: {0}", dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.Characters))
        Console.WriteLine("Characters with spaces: {0}", dc.Document.Properties.BuiltIn(BuiltInDocumentProperty.CharactersWithSpaces))
    End Sub
End Module