Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        AddText()
    End Sub

    ''' <summary>
    ''' How to create a simple document with text. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/text.php
    ''' </remarks>
    Sub AddText()
        Dim documentPath As String = "Text.docx"

        ' Let's create a new document.
        Dim dc As New DocumentCore()

        ' Create a new section and add into the document.
        Dim section As New Section(dc)
        dc.Sections.Add(section)

        ' Create a new paragraph and add into the section.
        Dim par As New Paragraph(dc)
        section.Blocks.Add(par)

        ' Create Inline-derived objects with text.
        Dim run1 As New Run(dc, "This is a rich")
        run1.CharacterFormat = New CharacterFormat() With {
            .FontName = "Times New Roman",
            .Size = 18.0,
            .FontColor = New Color(112, 173, 71)
        }

        Dim run2 As New Run(dc, " formatted text")
        run2.CharacterFormat = New CharacterFormat() With {
            .FontName = "Arial",
            .Size = 10.0,
            .FontColor = New Color("#0070C0")
        }

        Dim spch3 As New SpecialCharacter(dc, SpecialCharacterType.LineBreak)

        Dim run4 As New Run(dc, "with a line break.")
        run4.CharacterFormat = New CharacterFormat() With {
            .FontName = "Times New Roman",
            .Size = 10.0,
            .FontColor = Color.Black
        }

        ' Add our inlines into the paragraph.
        par.Inlines.Add(run1)
        par.Inlines.Add(run2)
        par.Inlines.Add(spch3)
        par.Inlines.Add(run4)

        ' Save our document into the DOCX format.
        dc.Save(documentPath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub
End Module