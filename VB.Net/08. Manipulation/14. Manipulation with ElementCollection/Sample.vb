Imports System.IO
Imports System.Linq
Imports SautinSoft.Document

Module Sample
    Sub Main()
        Manipulation()
    End Sub

    ''' <summary>
    ''' Manipulation with ElementCollection. Split 1st Paragraph by separate Runs and insert each Run into a new Paragraph.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/elementcollection-manipulation.php
    ''' </remarks>
    Sub Manipulation()
        Dim filePath As String = "..\..\..\example.docx"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)
        Dim filePathResult As String = "Result-file.pdf"
        Dim section As Section = dc.Sections(0)
        Dim paragraph As Paragraph = TryCast(section.Blocks(0), Paragraph)
        Dim i As Integer = 1
        Do While i < paragraph.Inlines.Count
            Dim inline As Inline = paragraph.Inlines(i)
            paragraph.Inlines.RemoveAt(1)
            section.Blocks.Add(New Paragraph(dc, inline))
        Loop
        dc.Save(filePathResult)
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePathResult) With {.UseShellExecute = True})
    End Sub
End Module