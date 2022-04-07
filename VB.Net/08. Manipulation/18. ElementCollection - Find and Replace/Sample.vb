Imports System
Imports SautinSoft.Document
Imports SautinSoft.Document.Tables

Module Sample
    Sub Main()
        FindAndReplace()
    End Sub

    ''' <summary>
    ''' Find an empty paragraphs in document, replace all tables into paragraphs.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/elementcollection-find-replace.php
    ''' </remarks>
    Sub FindAndReplace()
        Dim filePath As String = "..\..\..\example.docx"
        Dim result1 As String = "ResultEmptyParagraphs.docx"
        Dim result2 As String = "ResultReplacedTables.docx"

        Dim dc As DocumentCore = DocumentCore.Load(filePath)
        For Each par As Paragraph In dc.Sections(0).GetChildElements(False, ElementType.Paragraph)
            If par.Inlines.Count = 0 Then
                par.Inlines.Add(New Run(dc, "<empty paragraph>", New CharacterFormat() With {
                        .BackgroundColor = Color.Black,
                        .FontColor = Color.White
                    }))
            End If
        Next par
        dc.Save(result1)

        Dim i As Integer = 0
        Do While i < dc.Sections(0).Blocks.Count
            If TypeOf dc.Sections(0).Blocks(i) Is Table Then
                dc.Sections(0).Blocks(i) = New Paragraph(dc, New Run(dc, "HERE WAS THE TABLE", New CharacterFormat() With {.BackgroundColor = Color.Yellow}))
            End If
            i += 1
        Loop
        dc.Save(result2)

        ' Show the result.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(result1) With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(result2) With {.UseShellExecute = True})
    End Sub
End Module