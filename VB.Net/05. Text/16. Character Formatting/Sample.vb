Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        CharacterFormatting()
    End Sub

    ''' <summary>
    ''' This sample shows how to set character format. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/character-format.php
    ''' </remarks>
    Sub CharacterFormatting()
        Dim documentPath As String = "CharacterFormat.pdf"

        ' Let's create a simple document.
        Dim dc As New DocumentCore()

        ' Add a new section.
        dc.Sections.Add(New Section(dc))

        ' Add a paragraph.
        Dim p As New Paragraph(dc)
        p.ParagraphFormat.Alignment = HorizontalAlignment.Left
        dc.Sections(0).Blocks.Add(p)

        ' Create a formatted text (Run element) and add it into paragraph.
        Dim run1 As New Run(dc, "It's wide formatted text.")
        run1.CharacterFormat.AllCaps = True
        run1.CharacterFormat.BackgroundColor = Color.Pink
        run1.CharacterFormat.FontName = "Verdana"
        run1.CharacterFormat.Size = 26.0F
        run1.CharacterFormat.FontColor = New Color("#FFFFFF")

        p.Inlines.Add(run1)

        ' Create another Run element (container for characters).
        Dim run2 As New Run(dc, "Hi from SautinSoft!")
        run2.CharacterFormat.FontColor = Color.DarkGreen
        run2.CharacterFormat.UnderlineStyle = UnderlineType.Dashed
        run2.CharacterFormat.UnderlineColor = Color.Gray

        ' Add another formatted text into the paragraph.
        p.Inlines.Add(run2)

        ' Add new paragraph with formatted text.
        ' We are using ContentRange to insert text.
        dc.Content.Start.Insert("This is the first paragraph." & vbLf, New CharacterFormat() With {
                .FontName = "Calibri",
                .Size = 16.0,
                .FontColor = Color.Orange,
                .Bold = True
            })
        TryCast(dc.Sections(0).Blocks(0), Paragraph).ParagraphFormat.Alignment = HorizontalAlignment.Center

        dc.Content.End.Insert("Bold", New CharacterFormat() With {
                .Bold = True,
                .FontName = "Times New Roman",
                .Size = 11.0
            })
        dc.Content.End.Insert(" Italic ", New CharacterFormat() With {
                .Italic = True,
                .FontName = "Calibri",
                .Size = 11.0
            })
        dc.Content.End.Insert("Underline", New CharacterFormat() With {
                .UnderlineStyle = UnderlineType.Single,
                .FontName = "Calibri",
                .Size = 11.0
            })
        dc.Content.End.Insert(" ", New CharacterFormat() With {
                .Bold = True,
                .FontName = "Segoe UI",
                .Size = 11.0
            })
        dc.Content.End.Insert("Strikethrough", New CharacterFormat() With {
                .Strikethrough = True,
                .FontName = "Calibri",
                .Size = 11.0
            })

        ' Save our document into PDF format.
        dc.Save(documentPath, New PdfSaveOptions())

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub
End Module