Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        StyleInheritance()
    End Sub

    ''' <summary>
    ''' How the Styles Inheritance does work.
    ''' </summary>
    ''' <remarks>
    ''' https://sautinsoft.com/products/document/help/net/developer-guide/styles-inheritance.php
    ''' </remarks>
    Sub StyleInheritance()
        Dim docxPath As String = "StylesInheritance.docx"

        ' Let's create document.
        Dim dc As New DocumentCore()
        dc.DefaultCharacterFormat.FontColor = Color.Blue
        Dim section As New Section(dc)
        section.Blocks.Add(New Paragraph(dc, New Run(dc, "The document has Default Character Format with 'Blue' color.", New CharacterFormat() With {.Size = 18})))
        dc.Sections.Add(section)

        ' Create a new Paragraph and Style with 'Yellow' background.
        Dim par As New Paragraph(dc)
        Dim styleYellowBg As New ParagraphStyle("YellowBackground")
        styleYellowBg.CharacterFormat.BackgroundColor = Color.Yellow
        dc.Styles.Add(styleYellowBg)
        par.ParagraphFormat.Style = styleYellowBg

        par.Inlines.Add(New Run(dc, "This paragraph has Style 'Yellow Background' and it inherits 'Blue Color' from the document's DefaultCharacterFormat."))
        par.Inlines.Add(New SpecialCharacter(dc, SpecialCharacterType.LineBreak))
        Dim run1 As New Run(dc, "This Run doesn't have a style, but it inherits 'Yellow Background' from the paragraph style and 'Blue Color' from the document's DefaultCharacterFormat.")
        run1.CharacterFormat.Italic = True
        par.Inlines.Add(run1)
        par.Inlines.Add(New SpecialCharacter(dc, SpecialCharacterType.LineBreak))

        Dim run2 As New Run(dc, " This run has own Style with 'Green Color'.")
        Dim styleGreenText As New CharacterStyle("GreenText")
        styleGreenText.CharacterFormat.FontColor = Color.Green
        dc.Styles.Add(styleGreenText)
        run2.CharacterFormat.Style = styleGreenText
        par.Inlines.Add(run2)

        Dim par2 As New Paragraph(dc)
        Dim run3 As New Run(dc, "This is a new paragraph without a style. This is a Run also without style. " & "But they both inherit 'Blue Color' from their parent - the document.")
        par2.Inlines.Add(run3)
        section.Blocks.Add(par)
        section.Blocks.Add(par2)

        ' Save our document into DOCX format.
        dc.Save(docxPath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(docxPath) With {.UseShellExecute = True})
    End Sub

End Module