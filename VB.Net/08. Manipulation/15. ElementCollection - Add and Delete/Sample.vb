Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        AddAndDeleteParagraphs()
    End Sub

    ''' <summary>
    ''' ElementCollection: Adds 20 paragraphs into document and delete 10 of them.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/elementcollection-add-delete.php
    ''' </remarks>
    Sub AddAndDeleteParagraphs()
        Dim dc As New DocumentCore()
        Dim section As New Section(dc)
        dc.Sections.Add(section)
        For i As Integer = 0 To 19
            Dim par As New Paragraph(dc, "Text " & i.ToString())
            section.Blocks.Add(par)
        Next i
        dc.Save("ResultFull.docx")
        Dim j As Integer = 0
        Do While j < section.Blocks.Count
            section.Blocks.RemoveAt(j)
            j += 1
        Loop
        dc.Save("ResultShort.docx")
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo("ResultFull.docx") With {.UseShellExecute = True})
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo("ResultShort.docx") With {.UseShellExecute = True})
    End Sub
End Module