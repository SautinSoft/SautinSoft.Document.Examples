Imports System
Imports System.Text
Imports SautinSoft.Document
Imports SautinSoft.Document.CustomMarkups
Imports System.IO

Namespace Example
    Friend Class Program
        Shared Sub Main(ByVal args() As String)
            InsertPlainText()
        End Sub
        ''' <summary>
        ''' Inserting a plain text content control.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/content-controls-insert-plain-text-net-csharp-vb.php
        ''' </remarks>

        Private Shared Sub InsertPlainText()
            ' Let's create a simple document.
            Dim dc As New DocumentCore()

            ' Create a plain text content control.
            Dim pt As New BlockContentControl(dc, ContentControlType.PlainText)

            ' Add a new section.
            dc.Sections.Add(New Section(dc, pt))

            ' Add the content control properties.
            pt.Properties.Title = "Title"
            pt.Properties.Multiline = True
            pt.Properties.Color = Color.Blue
            pt.Document.DefaultCharacterFormat.FontColor = Color.Orange

            ' Add new paragraph with formatted text.
            pt.Blocks.Add(New Paragraph(dc, New Run(dc, "This is first paragraph with symbols added on a new line."),
            New SpecialCharacter(dc, SpecialCharacterType.LineBreak),
            New Run(dc, "This is a new line in the first paragraph."),
            New SpecialCharacter(dc, SpecialCharacterType.LineBreak),
            New Run(dc, "Insert the ""Wingdings"" font family with formatting."),
            New Run(dc, ChrW(&HFC).ToString() & ChrW(&HF0).ToString() & ChrW(&H32).ToString(),
            New CharacterFormat() With
                 {
                    .FontName = "Wingdings",
                    .FontColor = New Color("#000000"),
                    .Size = 48
                })))

            ' Save our document into DOCX format.
            Dim resultPath As String = "result.docx"
            dc.Save(resultPath, New DocxSaveOptions())

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
