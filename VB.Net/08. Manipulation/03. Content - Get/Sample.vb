Imports System
Imports SautinSoft.Document
Imports System.Text

Module Sample
    Sub Main()
        GetContent()
    End Sub

    ''' <summary>
    ''' How to get a content from a document.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/get-content-net-csharp-vb.php
    ''' </remarks>
    Sub GetContent()
        ' Path to an input document.
        Dim documentPath As String = "..\..\..\example.docx"

        Dim dc As DocumentCore = DocumentCore.Load(documentPath)

        Dim sb As New StringBuilder()

        ' Get content of each paragraph in the document.
        For Each par As Paragraph In dc.GetChildElements(True, ElementType.Paragraph)
            ' The property 'Content' returns the content as ContentRange.
            ' Get content and append it into StringBuilder.
            sb.AppendFormat("Paragraph: {0}", par.Content.ToString())
            sb.AppendLine()
        Next par

        ' Get content of each Run where the text color is Red.
        For Each run As Run In dc.GetChildElements(True, ElementType.Run)
            If run.CharacterFormat.FontColor = Color.Red Then
                ' The property 'Content' returns the content as ContentRange.
                ' Get content and append it into StringBuilder.
                sb.AppendFormat("Red color: {0}", run.Content.ToString())
                sb.AppendLine()
            End If
        Next run
        Console.WriteLine(sb.ToString())
        Console.ReadKey()
    End Sub
End Module