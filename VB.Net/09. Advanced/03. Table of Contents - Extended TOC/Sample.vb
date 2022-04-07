Option Infer On

Imports System
Imports System.Linq
Imports SautinSoft.Document

Module Sample
    Sub Main()
        ExtendedTOC()
    End Sub

    ''' <summary>
    ''' Create extended table of contents in word document.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-extended-table-of-contents-in-word-document-net-csharp-vb.php
    ''' </remarks>
    Sub ExtendedTOC()
        Dim resultFile As String = "Extended-Table-Of-Contents.docx"
        ' First of all, create an instance of DocumentCore.
        Dim dc As New DocumentCore()

        ' Create and add Heading1 style. For "Chapter 1" and "Chapter 2".
        Dim Heading1Style As ParagraphStyle = CType(Style.CreateStyle(StyleTemplateType.Heading1, dc), ParagraphStyle)
        Heading1Style.ParagraphFormat.LineSpacing = 3
        Heading1Style.CharacterFormat.Size = 18
        ' #358CCB - blue
        Heading1Style.CharacterFormat.FontColor = New Color("#358CCB")
        dc.Styles.Add(Heading1Style)

        ' Create and add Heading2 style. For "SupChapter 1-1" and "SubChapter 2-1".
        Dim Heading2Style As ParagraphStyle = CType(Style.CreateStyle(StyleTemplateType.Heading2, dc), ParagraphStyle)
        Heading2Style.ParagraphFormat.LineSpacing = 2
        Heading2Style.CharacterFormat.Size = 14
        ' #FF9900 - orange
        Heading2Style.CharacterFormat.FontColor = New Color("#FF9900")
        dc.Styles.Add(Heading2Style)

        ' Create and add TOC style.
        Dim TOCStyle As ParagraphStyle = CType(Style.CreateStyle(StyleTemplateType.Subtitle, dc), ParagraphStyle)
        TOCStyle.ParagraphFormat.OutlineLevel = OutlineLevel.BodyText
        TOCStyle.ParagraphFormat.Alignment = HorizontalAlignment.Center
        TOCStyle.CharacterFormat.Bold = True
        ' #358CCB - blue
        TOCStyle.CharacterFormat.FontColor = New Color("#358CCB")
        dc.Styles.Add(TOCStyle)

        ' Add new section.
        Dim section As New Section(dc)
        dc.Sections.Add(section)

        ' Add TOC Header.
        section.Blocks.Add(New Paragraph(dc, "Table of Contents") With {
            .ParagraphFormat = New ParagraphFormat With {.Style = TOCStyle}
        })

        ' Create and add TOC (Table of Contents).
        section.Blocks.Add(New TableOfEntries(dc, FieldType.TOC))

        ' Add TOC Ending.
        section.Blocks.Add(New Paragraph(dc, "The End") With {
            .ParagraphFormat = New ParagraphFormat With {
                .Alignment = HorizontalAlignment.Center,
                .BackgroundColor = Color.Gray
            }
        })

        ' Add the document content (Chapter 1).
        ' Add Chapter 1.
        section.Blocks.Add(New Paragraph(dc, "Chapter 1") With {
            .ParagraphFormat = New ParagraphFormat With {
                .Style = Heading1Style,
                .PageBreakBefore = True
            }
        })

        ' Add SubChapter 1-1.
        section.Blocks.Add(New Paragraph(dc, String.Format("Subchapter 1-1")) With {
            .ParagraphFormat = New ParagraphFormat With {.Style = Heading2Style}
        })

        ' Add the content of Chapter 1 / Subchapter 1-1.
        section.Blocks.Add(New Paragraph(dc, "«Document.Net» will help you in development of applications which operates with DOCX, RTF, PDF, HTML and Text documents.After adding of the reference to(SautinSoft.Document.dll) - it's 100% C# managed assembly you will be able to create a new document, parse an existing, modify anything what you want.") With {
            .ParagraphFormat = New ParagraphFormat With {
                .LeftIndentation = 10,
                .RightIndentation = 10,
                .SpecialIndentation = 20,
                .LineSpacing = 20,
                .LineSpacingRule = LineSpacingRule.Exactly,
                .SpaceBefore = 20,
                .SpaceAfter = 20
            }
        })

        ' Let's add another page break into.
        section.Blocks.Add(New Paragraph(dc, New SpecialCharacter(dc, SpecialCharacterType.PageBreak)))

        ' Add the document content (Chapter 2).
        ' Add Chapter 2.
        section.Blocks.Add(New Paragraph(dc, "Chapter 2") With {
            .ParagraphFormat = New ParagraphFormat With
            {.Style = Heading1Style}
        })

        ' Add SubChapter 2-1.
        section.Blocks.Add(New Paragraph(dc, String.Format("Subchapter 2-1")) With {
            .ParagraphFormat = New ParagraphFormat With {.Style = Heading2Style}
        })

        ' Add the content of Chapter 2 / Subchapter 2-1.
        section.Blocks.Add(New Paragraph(dc, "Requires only .Net 4.0 or above. Our product is compatible with all .Net languages and supports all Operating Systems where .Net Framework can be used. Note that «Document .Net» is entirely written in managed C#, which makes it absolutely standalone and an independent library. Of course, No dependency on Microsoft Word.") With {
            .ParagraphFormat = New ParagraphFormat With {
                .LeftIndentation = 10,
                .RightIndentation = 10,
                .SpecialIndentation = 20,
                .LineSpacing = 20,
                .LineSpacingRule = LineSpacingRule.Exactly,
                .SpaceBefore = 20,
                .SpaceAfter = 20
            }
        })

        ' Update TOC (TOC can be updated only after all document content is added).
        Dim tableofcontents = CType(dc.GetChildElements(True, ElementType.TableOfEntries).FirstOrDefault(), TableOfEntries)
        tableofcontents.Update()

        ' Apply the style for the TOC.
        For Each par As Paragraph In tableofcontents.Entries
            par.ParagraphFormat.Style = TOCStyle
        Next par

        ' Update TOC's page numbers.
        ' Page numbers are automatically updated in that case.
        dc.GetPaginator(New PaginatorOptions() With {.UpdateFields = True})

        ' Save the document as DOCX file.
        dc.Save(resultFile)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultFile) With {.UseShellExecute = True})
    End Sub

End Module