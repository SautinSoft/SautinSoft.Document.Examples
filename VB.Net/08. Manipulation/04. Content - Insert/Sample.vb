Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document

Module Sample
    Sub Main()
        InsertContent()
    End Sub

    ''' <summary>
    ''' Creates a new document with some content.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-content-net-csharp-vb.php
    ''' </remarks>
    Sub InsertContent()
        Dim documentPath As String = "InsertContent.pdf"

        ' Let's create a new document.
        Dim dc As New DocumentCore()

        ' Add new section.
        Dim section As New Section(dc)
        dc.Sections.Add(section)

        ' You may add a new paragraph using a classic way:
        section.Blocks.Add(New Paragraph(dc, "This is a first line in 1st paragraph!"))
        Dim par1 As Paragraph = TryCast(section.Blocks(0), Paragraph)
        par1.ParagraphFormat.Alignment = HorizontalAlignment.Center

        ' But here, let's see how to do it using ContentRange:

        ' 1. Insert a content as Text.
        dc.Content.End.Insert(vbLf & "This is the first line in 2nd paragraph.", New CharacterFormat() With {
            .Size = 25,
            .FontColor = Color.Blue,
            .Bold = True
        })

        ' 2. Insert a content as HTML (at the start).
        dc.Content.Start.Insert("Hello from HTML!", New HtmlLoadOptions())

        ' 3. Insert a content as RTF.
        dc.Content.End.Insert("{\rtf1 \b The line from RTF\b0!\par}", New RtfLoadOptions())

        ' 4. Insert a content of SpecialCharacter.
        dc.Content.End.Insert((New SpecialCharacter(dc, SpecialCharacterType.LineBreak)).Content)

        ' 5. Insert a content as ContentRange.
        ' Find first content of "HTML" and insert it at the end.
        Dim cr As ContentRange = dc.Content.Find("HTML").First()
        dc.Content.End.Insert(cr)

        ' 6. Insert the content of Paragraph at the end.
        Dim p As New Paragraph(dc)
        Dim run As New Run(dc, "As summarize, you can insert content of any class " & "derived from Element: Table, Shape, Paragraph, Run, HeaderFooter and even a whole DocumentCore!")
        run.CharacterFormat.FontColor = Color.DarkGreen
        p.Inlines.Add(run)
        dc.Content.End.Insert(p.Content)

        ' Save our document into PDF format.
        dc.Save(documentPath, New PdfSaveOptions() With {.Compliance = PdfCompliance.PDF_A1a})

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub
End Module