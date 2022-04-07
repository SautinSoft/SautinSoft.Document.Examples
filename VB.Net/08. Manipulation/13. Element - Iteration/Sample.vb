Imports System
Imports SautinSoft.Document
Imports System.IO
Imports System.Linq
Imports System.Text

Module Sample
    Sub Main()
        IterationElement()
    End Sub

    ''' <summary>
    ''' Calculate sections, paragraphs, inlines, runs and fields in DOCX document.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/iteration-in-element-collection-net-csharp-vb.php
    ''' </remarks>
    Sub IterationElement()
        Dim dc As DocumentCore = DocumentCore.Load("..\..\..\Parsing.docx", LoadOptions.DocxDefault)
        Dim numberOfSections As Integer = dc.Sections.Count
        Dim numberOfParagraphs As Integer = dc.GetChildElements(True, ElementType.Paragraph).Count()
        Dim numberOfRunsAndFields As Integer = dc.GetChildElements(True, ElementType.Run, ElementType.Field).Count()
        Dim numberOfInlines As Integer = dc.GetChildElements(True).OfType(Of Inline)().Count()
        Dim elements As Integer = dc.Sections(0).GetChildElements(True).Count()
        Dim sb As New StringBuilder()
        sb.AppendLine("File has:")
        sb.AppendLine(numberOfSections & " section")
        sb.AppendLine(numberOfParagraphs & " paragraphs")
        sb.AppendLine(numberOfRunsAndFields & " runs and fields")
        sb.AppendLine(numberOfInlines & " inlines")
        sb.AppendLine("First section contains " & elements & " elements")
        Console.WriteLine(sb.ToString())
        Console.ReadKey()
    End Sub
End Module