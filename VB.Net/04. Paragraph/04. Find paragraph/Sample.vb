Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document

Namespace Example
    Friend Class Program
        Shared Sub Main(ByVal args() As String)
            FindParagraph()
        End Sub
        ''' <summary>
        ''' Find all paragraphs aligned by center in DOCX document and mark it by yellow.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/find-paragraphs-in-docx-document-net-csharp-vb.php
        ''' </remarks>
        Private Shared Sub FindParagraph()
            Dim filePath As String = "..\..\..\example.docx"
            Dim fileResult As String = "Result.docx"
            Dim dc As DocumentCore = DocumentCore.Load(filePath)

            For Each par As Paragraph In dc.GetChildElements(True, ElementType.Paragraph).Where(Function(p) (TryCast(p, Paragraph)).ParagraphFormat.Alignment = HorizontalAlignment.Center)
                par.ParagraphFormat.BackgroundColor = Color.Yellow
            Next par
            dc.Save(fileResult)
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(fileResult) With {.UseShellExecute = True})

        End Sub
    End Class
End Namespace
