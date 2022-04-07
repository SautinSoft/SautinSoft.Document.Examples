Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        ParagraphFormatting()
    End Sub

    ''' <summary>
    ''' This sample shows how to specify a paragraph formatting. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/paragraph-format.php
    ''' </remarks>
    Sub ParagraphFormatting()
        Dim documentPath As String = "ParagraphFormatting.docx"

        ' Let's create a simple document.
        Dim dc As New DocumentCore()

        ' Add new section
        Dim s As New Section(dc)
        dc.Sections.Add(s)

        ' Paragraph 1.
        Dim p As New Paragraph(dc, "First paragraph!")
        dc.Sections(0).Blocks.Add(p)

        p.ParagraphFormat.Alignment = HorizontalAlignment.Left
        p.ParagraphFormat.SpaceAfter = 20.0
        p.ParagraphFormat.Borders.Add(MultipleBorderTypes.All, BorderStyle.Single, Color.Orange, 2.0F)

        ' Paragraph 2.
        dc.Content.End.Insert(vbLf & "This is a second paragraph.", New CharacterFormat() With {
            .FontName = "Calibri",
            .Size = 16.0,
            .FontColor = Color.Black,
            .Bold = True
        })
        TryCast(dc.Sections(0).Blocks(1), Paragraph).ParagraphFormat.Alignment = HorizontalAlignment.Center

        ' Create multiple paragraphs.
        For i As Integer = 0 To 9
            Dim pN As New Paragraph(dc, String.Format("Paragraph {0}. ", i + 1))
            dc.Sections(0).Blocks.Add(pN)

            pN.Content.End.Insert((New SpecialCharacter(dc, SpecialCharacterType.Tab)).Content)
            Dim run As New Run(dc, "Hello!")
            run.CharacterFormat.FontColor = Color.White
            pN.Content.End.Insert(run.Content)

            pN.ParagraphFormat.BackgroundColor = New Color(CInt(&HFF358CCB * (i + 1)))
            pN.ParagraphFormat.SpaceBefore = LengthUnitConverter.Convert(1, LengthUnit.Centimeter, LengthUnit.Point)
            pN.ParagraphFormat.SpaceAfter = LengthUnitConverter.Convert(10, LengthUnit.Millimeter, LengthUnit.Point)
        Next i

        ' Save our document into DOCX format.
        dc.Save(documentPath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub

End Module