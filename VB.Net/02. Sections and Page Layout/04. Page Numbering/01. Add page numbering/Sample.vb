Imports System
Imports System.IO
Imports System.Linq
Imports SautinSoft.Document

Module Sample
    Sub Main()
        PageNumbering()
    End Sub

    ''' <summary>
    ''' Creates a new document with page numbering: Page N of M.
    ''' </summary>
    ''' <remarks>
    ''' https://sautinsoft.com/products/document/help/net/developer-guide/page-numbering.php
    ''' </remarks>
    Sub PageNumbering()
        Dim documentPath As String = "PageNumbering.docx"

        ' Let's create a new document with multiple pages.
        Dim dc As New DocumentCore()

        Dim pagesText() As String = {"One", "Two", "Three", "Four", "Five"}
        Dim r As New Random()

        ' Create a new section.
        Dim section As New Section(dc)
        dc.Sections.Add(section)

        ' We place our page numbers into the footer.
        ' Therefore we've to create a footer.
        Dim footer As New HeaderFooter(dc, HeaderFooterType.FooterDefault)

        ' Create a new paragraph to insert a page numbering.
        ' So that, our page numbering looks as: Page N of M.
        Dim par As New Paragraph(dc)
        par.ParagraphFormat.Alignment = HorizontalAlignment.Left
        Dim cf As New CharacterFormat() With {
                .FontName = "Arial",
                .Size = 12.0
            }
        par.Content.Start.Insert("Page ", cf.Clone())

        ' Page numbering is a Field.
        ' Create two fields: FieldType.Page and FieldType.NumPages.
        Dim fPage As New Field(dc, FieldType.Page)
        fPage.CharacterFormat = cf.Clone()
        par.Content.End.Insert(fPage.Content)
        par.Content.End.Insert(" of ", cf.Clone())
        Dim fPages As New Field(dc, FieldType.NumPages)
        fPages.CharacterFormat = cf.Clone()
        par.Content.End.Insert(fPages.Content)

        ' Add the paragraph with Fields into the footer.
        footer.Blocks.Add(par)

        ' Add the footer into the section.
        section.HeadersFooters.Add(footer)

        ' Add some paragraphs with page breaks in the document,
        ' to make several pages.
        For Each text As String In pagesText
            Dim p As New Paragraph(dc)
            p.ParagraphFormat.Alignment = HorizontalAlignment.Center
            section.Blocks.Add(p)

            Dim color As String = String.Format("#{0:X2}{1:X2}{2:X2}", r.Next(0, 255), r.Next(0, 255), r.Next(0, 255))

            p.Content.Start.Insert(text, New CharacterFormat() With {
                    .FontName = "Arial",
                    .Size = 72.0,
                    .FontColor = New Color(color)
                })
            If (text <> pagesText.Last()) Then
                p.Content.End.Insert((New SpecialCharacter(dc, SpecialCharacterType.PageBreak)).Content)
            End If
        Next text

        ' Save our document into DOCX format.
        dc.Save(documentPath, New DocxSaveOptions())

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub
End Module