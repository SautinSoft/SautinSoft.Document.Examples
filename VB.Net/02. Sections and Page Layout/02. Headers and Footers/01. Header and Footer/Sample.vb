Imports System
Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Module Sample
    Sub Main()
        HeadersAndFooters()
    End Sub

    ''' <summary>
    ''' How to add a header and footer into a document.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/headersfooters.php
    ''' </remarks>
    Sub HeadersAndFooters()
        Dim documentPath As String = "HeadersAndFooters.docx"

        ' Let's create a simple document.
        Dim dc As New DocumentCore()

        ' Add a new section in the document.
        Dim s As New Section(dc)
        dc.Sections.Add(s)

        ' Let's add a paragraph with text.
        Dim p As New Paragraph(dc)
        dc.Sections(0).Blocks.Add(p)

        p.ParagraphFormat.Alignment = HorizontalAlignment.Justify
        p.Content.Start.Insert("Once upon a time, in a far away swamp, there lived an ogre named Shrek whose precious " & "solitude is suddenly shattered by an invasion of annoying fairy tale characters...", New CharacterFormat() With {
            .Size = 12,
            .FontName = "Arial"
        })

        ' Create a new header with formatted text.
        Dim header As New HeaderFooter(dc, HeaderFooterType.HeaderDefault)
        header.Content.Start.Insert("Shrek and Donkey travel to the castle and split up to find Fiona.", New CharacterFormat() With {.Size = 14.0})

        ' Add the header into HeadersFooters collection of the 1st section.
        s.HeadersFooters.Add(header)

        ' Create a new footer with some text and image.
        Dim footer As New HeaderFooter(dc, HeaderFooterType.FooterDefault)

        ' Create a paragraph to insert it into the footer.
        Dim par As New Paragraph(dc)
        par.Content.Start.Insert("Shrek and Donkey travel to the castle and split up to find Fiona. ", New CharacterFormat() With {.Size = 14.0})
        par.ParagraphFormat.Alignment = HorizontalAlignment.Left

        ' Insert image into the paragraph.
        Dim wPt As Double = LengthUnitConverter.Convert(7, LengthUnit.Centimeter, LengthUnit.Point)
        Dim hPt As Double = LengthUnitConverter.Convert(7, LengthUnit.Centimeter, LengthUnit.Point)

        Dim pict As New Picture(dc, Layout.Inline(New Size(wPt, hPt)), "..\..\..\image1.jpg")
        par.Inlines.Add(pict)

        ' Add the paragraph into the Blocks collection of the footer.
        footer.Blocks.Add(par)

        ' Finally, add the footer into 1st section (HeadersFooters collection).
        s.HeadersFooters.Add(footer)

        ' Save the document into DOCX format.
        dc.Save(documentPath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub
End Module