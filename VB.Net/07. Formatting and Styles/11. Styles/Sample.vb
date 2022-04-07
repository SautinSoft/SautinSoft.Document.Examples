Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        Styles()
    End Sub

    ''' <summary>
    ''' This sample shows how to work with styles. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/styles.php
    ''' </remarks>
    Sub Styles()
        Dim docxPath As String = "Styles.docx"

        ' Let's create document.
        Dim dc As New DocumentCore()

        ' Create custom styles.
        Dim characterStyle As New CharacterStyle("CharacterStyle1")
        characterStyle.CharacterFormat.FontName = "Arial"
        characterStyle.CharacterFormat.UnderlineStyle = UnderlineType.Wave
        characterStyle.CharacterFormat.Size = 18

        Dim paragraphStyle As New ParagraphStyle("ParagraphStyle1")
        paragraphStyle.CharacterFormat.FontName = "Times New Roman"
        paragraphStyle.CharacterFormat.Size = 14
        paragraphStyle.ParagraphFormat.Alignment = HorizontalAlignment.Center

        ' Add styles to the document, then use it.
        dc.Styles.Add(characterStyle)
        dc.Styles.Add(paragraphStyle)


        ' Add text content.
        Dim par As New Paragraph(dc)
        par.ParagraphFormat.Style = paragraphStyle
        dc.Sections.Add(New Section(dc, par))
        par.Inlines.Add(New Run(dc, "Once upon a time, in a far away swamp, there lived an ogre named "))

        Dim run As New Run(dc, "Shrek")
        run.CharacterFormat.Style = characterStyle
        par.Inlines.Add(run)

        par.Inlines.Add(New Run(dc, " whose precious solitude is suddenly shattered by an invasion of annoying fairy tale characters..."))

        ' Save our document into DOCX format.
        dc.Save(docxPath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(docxPath) With {.UseShellExecute = True})
    End Sub
End Module