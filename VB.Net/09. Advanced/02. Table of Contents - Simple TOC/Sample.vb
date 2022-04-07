Option Infer On

Imports System
Imports System.Linq
Imports SautinSoft.Document

Module Sample
    Sub Main()
        TOC()
    End Sub

    ''' <summary>
    ''' Create a document with table of content.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/table-of-contents-toc.php
    ''' </remarks>
    Sub TOC()
        Dim resultFile As String = "Table-Of-Contents.docx"

        ' Let's create a simple document.
        Dim dc As New DocumentCore()

        ' Create and add Heading 1 style for TOC.
        Dim Heading1Style As ParagraphStyle = CType(Style.CreateStyle(StyleTemplateType.Heading1, dc), ParagraphStyle)
        dc.Styles.Add(Heading1Style)

        ' Create and add Heading 2 style for TOC.
        Dim Heading2Style As ParagraphStyle = CType(Style.CreateStyle(StyleTemplateType.Heading2, dc), ParagraphStyle)
        dc.Styles.Add(Heading2Style)

        ' Add new section.
        Dim section As New Section(dc)
        dc.Sections.Add(section)

        ' Add TOC title in the DOCX document.
        section.Blocks.Add(New Paragraph(dc, "Table of Contents"))

        ' Create and add new TOC.
        section.Blocks.Add(New TableOfEntries(dc, FieldType.TOC))

        section.Blocks.Add(New Paragraph(dc, "The end."))

        ' Let's add a page break into our paragraph.
        section.Blocks.Add(New Paragraph(dc, New SpecialCharacter(dc, SpecialCharacterType.PageBreak)))

        ' Add document content.
        ' Add Chapter 1
        section.Blocks.Add(New Paragraph(dc, "Chapter 1") With {
            .ParagraphFormat = New ParagraphFormat With
            {.Style = Heading1Style}
        })

        ' Add SubChapter 1-1
        section.Blocks.Add(New Paragraph(dc, String.Format("Subchapter 1-1")) With {
            .ParagraphFormat = New ParagraphFormat With
             {.Style = Heading2Style}
        })
        ' Add the content of Chapter 1 / Subchapter 1-1
        section.Blocks.Add(New Paragraph(dc, "«Document .Net» will help you in development of applications which operates with DOCX, RTF, PDF and Text documents. After adding of the reference to (SautinSoft.Document.dll) - it's 100% C# managed assembly you will be able to create a new document, parse an existing, modify anything what you want."))

        ' Let's add an another page break into our paragraph.
        section.Blocks.Add(New Paragraph(dc, New SpecialCharacter(dc, SpecialCharacterType.PageBreak)))

        ' Add document content.
        ' Add Chapter 2
        section.Blocks.Add(New Paragraph(dc, "Chapter 2") With {
            .ParagraphFormat = New ParagraphFormat With
            {.Style = Heading1Style}
        })

        ' Add SubChapter 2-1
        section.Blocks.Add(New Paragraph(dc, String.Format("Subchapter 2-1")) With {
            .ParagraphFormat = New ParagraphFormat With
            {.Style = Heading2Style}
        })

        ' Add the content of Chapter 2 / Subchapter 2-1
        section.Blocks.Add(New Paragraph(dc, "Requires only .Net 4.0 or above. Our product is compatible with all .Net languages and supports all Operating Systems where .Net Framework can be used. Note that «Document .Net» is entirely written in managed C#, which makes it absolutely standalone and an independent library. Of course, No dependency on Microsoft Word."))

        ' Update TOC (TOC can be updated only after all document content is added).
        'INSTANT VB NOTE: The variable toc was renamed since Visual Basic does not handle local variables named the same as class members well:
        Dim toc_Renamed = CType(dc.GetChildElements(True, ElementType.TableOfEntries).FirstOrDefault(), TableOfEntries)
        toc_Renamed.Update()

        ' Update TOC's page numbers.
        ' Page numbers are automatically updated in that case.
        dc.GetPaginator(New PaginatorOptions() With {.UpdateFields = True})

        ' Save DOCX to a file
        dc.Save(resultFile)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultFile) With {.UseShellExecute = True})
    End Sub
End Module