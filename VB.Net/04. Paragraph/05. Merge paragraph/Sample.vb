Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document

Namespace Example
    Friend Class Program
        Shared Sub Main(ByVal args() As String)
            MergeParagraphs()
        End Sub
        ''' <summary>
        ''' Merge all paragraphs into a single in an existing PDF document.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/merge-paragraphs-in-pdf-document-net-csharp-vb.php
        ''' </remarks>
        Private Shared Sub MergeParagraphs()
            Dim inpFile As String = "..\..\..\example.pdf"
            Dim outFile As String = "Result.pdf"
            Dim dc As DocumentCore = DocumentCore.Load(inpFile)

            Dim firstPar As Paragraph = TryCast(dc.GetChildElements(True, ElementType.Paragraph).First(), Paragraph)

            Dim lastIndex As Integer = firstPar.Inlines.Count

            For Each par As Paragraph In dc.GetChildElements(True, ElementType.Paragraph).Reverse().Where(Function(p) p IsNot firstPar)
                Dim last As Integer = lastIndex
                For Each inline As Inline In par.Inlines
                    firstPar.Inlines.Insert(last, inline.Clone(True))
                    last += 1
                Next inline
                par.Content.Delete()
            Next par

            dc.Save(outFile)
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(inpFile) With {.UseShellExecute = True})
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
