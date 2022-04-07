Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        MultilevelLists()
    End Sub

    ''' <summary>
    ''' How to create multilevel ordered and unordered lists.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-multilevel-list-in-docx-document-net-csharp-vb.php
    ''' </remarks>
    Sub MultilevelLists()
        Dim documentPath As String = "MultilvelLists.docx"

        ' Let's create a new document.
        Dim dc As New DocumentCore()

        ' Add a new section.
        Dim s As New Section(dc)
        dc.Sections.Add(s)

        Dim myCollection() As String = {"One", "Two", "Three", "Four", "Five"}

        ' Create list style.
        Dim ls As New ListStyle("MyListDot", ListTemplateType.NumberWithDot)
        dc.Styles.Add(ls)

        ' Add the collection of paragraphs marked as ordered list.
        Dim level As Integer = 0
        For Each listText As String In myCollection
            Dim p As New Paragraph(dc)
            dc.Sections(0).Blocks.Add(p)

            p.Content.End.Insert(listText, New CharacterFormat() With {
                .Size = 14.0,
                .FontColor = Color.Black
            })
            p.ListFormat.Style = ls
            p.ListFormat.ListLevelNumber = level
            level += 1
            p.ParagraphFormat.SpaceAfter = 0
        Next listText

        ' Add the collection of paragraphs marked as unordered list (bullets).
        ' Create list style.
        Dim ls1 As New ListStyle("MyListBullet", ListTemplateType.Bullet)
        dc.Styles.Add(ls1)

        level = 0
        For Each listText As String In myCollection
            Dim p As New Paragraph(dc)
            dc.Sections(0).Blocks.Add(p)

            p.Content.End.Insert(listText, New CharacterFormat() With {
                .Size = 14.0,
                .FontColor = Color.Black
            })
            p.ListFormat.Style = ls1
            p.ListFormat.ListLevelNumber = level
            level += 1
            p.ParagraphFormat.SpaceAfter = 0
        Next listText

        ' Save our document into DOCX file.
        dc.Save(documentPath, New DocxSaveOptions())

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub

End Module