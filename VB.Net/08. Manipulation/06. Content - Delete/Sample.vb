Option Infer On
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document

Module Sample
    Sub Main()
        DeleteContent()
    End Sub

    ''' <summary>
    ''' Open a document and delete some content.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/delete-content-net-csharp-vb.php
    ''' </remarks>
    Sub DeleteContent()
        Dim loadPath As String = "..\..\..\example.docx"
        Dim savePath As String = "Result.docx"

        Dim dc As DocumentCore = DocumentCore.Load(loadPath)

        ' Remove the text "This" from all paragraphs in 1st section.
        For Each par As Paragraph In dc.Sections(0).GetChildElements(True, ElementType.Paragraph)
            Dim findText = par.Content.Find("This")

            If findText IsNot Nothing Then
                For Each cr As ContentRange In findText
                    cr.Delete()
                Next cr
            End If
        Next par
        dc.Save(savePath)
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(loadPath) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(savePath) With {.UseShellExecute = True})
    End Sub

End Module