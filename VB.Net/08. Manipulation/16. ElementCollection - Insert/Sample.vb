Imports System
Imports SautinSoft.Document
Imports SautinSoft.Document.Tables

Module Sample
    Sub Main()
        InsertParagraphCount()
    End Sub

    ''' <summary>
    ''' Inserts a new Run (Text element) at the start of each paragraph.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/elementcollection-insert.php
    ''' </remarks>
    Sub InsertParagraphCount()
        Dim filePath As String = "..\..\..\example.docx"
        Dim dc As DocumentCore = DocumentCore.Load(filePath)
        Dim paragraphNum As Integer = 1
        For Each el As Element In dc.Sections(0).GetChildElements(False)
            If TypeOf el Is Paragraph Then
                ' Insert a new Run into Paragraph.InlineCollection 'Inlines'.
                ' InlineCollection is descendant of the base abstract class ElementCollection.
                TryCast(el, Paragraph).Inlines.Insert(0, New Run(dc, "Paragraph " & paragraphNum.ToString() & " - ", New CharacterFormat() With {
                        .BackgroundColor = Color.Orange,
                        .FontColor = Color.White
                    }))
                paragraphNum += 1
            End If
        Next el
        dc.Save("Result.docx")

        ' Show the result.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo("Result.docx") With {.UseShellExecute = True})
    End Sub
End Module