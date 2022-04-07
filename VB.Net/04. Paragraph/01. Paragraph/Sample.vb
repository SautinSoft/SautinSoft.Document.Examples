Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        Paragraphs()
    End Sub

    ''' <summary>
    ''' Creates a new document with a paragraph.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/paragraph.php
    ''' </remarks>
    Public Sub Paragraphs()
        Dim dc As New DocumentCore()

        Dim filePath As String = "Result-file.docx"

        dc.Sections.Add(New Section(dc, New Paragraph(dc, "Text is right aligned.") With {
                .ParagraphFormat = New ParagraphFormat With {.Alignment = HorizontalAlignment.Right}
            },
            New Paragraph(dc, "This paragraph has the following properties: Left indentation is 15 points, right indentation is 5 centimeters, hanging indentation is 10 points, line spacing is exactly 14 points, space before and space after are 10 points.") With {
                .ParagraphFormat = New ParagraphFormat With {
                    .LeftIndentation = 15,
                    .RightIndentation = LengthUnitConverter.Convert(5, LengthUnit.Centimeter, LengthUnit.Point),
                    .SpecialIndentation = 10,
                    .LineSpacing = 14,
                    .LineSpacingRule = LineSpacingRule.Exactly,
                    .SpaceBefore = 10,
                    .SpaceAfter = 10
                }
            }))

        dc.Save(filePath)
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub
End Module