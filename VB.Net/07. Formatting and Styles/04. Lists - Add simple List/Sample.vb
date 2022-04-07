Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        SimpleLists()
    End Sub

    ''' <summary>
    ''' How to create a simple document with ordered and unordered lists. 
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-simple-list-in-pdf-document-net-csharp-vb.php
    ''' </remarks>
    Sub SimpleLists()
        Dim documentPath As String = "Lists.pdf"

        ' Let's create a simple document.
        Dim dc As New DocumentCore()

        ' Add a new section.
        Dim s As New Section(dc)
        dc.Sections.Add(s)

        Dim myCollection() As String = {"One", "Two", "Three", "Four", "Five"}

        ' Create new ordered list style.
        Dim ordList As New ListStyle("Simple Numbers", ListTemplateType.NumberWithDot)
        For Each f As ListLevelFormat In ordList.ListLevelFormats
            f.CharacterFormat.Size = 14
        Next f
        dc.Styles.Add(ordList)

        ' Add caption
        dc.Content.Start.Insert("This is a simple ordered list:", New CharacterFormat() With {
            .Size = 14.0,
            .FontColor = Color.Black
        })

        ' Add the collection of paragraphs marked as ordered list.
        Dim level As Integer = 0
        For Each listText As String In myCollection
            Dim p As New Paragraph(dc)
            dc.Sections(0).Blocks.Add(p)

            p.Content.End.Insert(listText, New CharacterFormat() With {
                .Size = 14.0,
                .FontColor = Color.Black
            })
            p.ListFormat.Style = ordList
            p.ListFormat.ListLevelNumber = level
            p.ParagraphFormat.SpaceAfter = 0
        Next listText

        ' Add the collection of paragraphs marked as unordered list (bullets).

        ' Add caption
        dc.Content.End.Insert(vbLf & vbLf & "This is a simple bulleted list:", New CharacterFormat() With {
            .Size = 14.0,
            .FontColor = Color.Black
        })

        ' Create list style.
        Dim bullList As New ListStyle("Bullets", ListTemplateType.Bullet)
        dc.Styles.Add(bullList)

        level = 0
        For Each listText As String In myCollection
            Dim p As New Paragraph(dc)
            dc.Sections(0).Blocks.Add(p)

            p.Content.End.Insert(listText, New CharacterFormat() With {
                .Size = 14.0,
                .FontColor = Color.Black
            })
            p.ListFormat.Style = bullList
            p.ListFormat.ListLevelNumber = level
            p.ParagraphFormat.SpaceAfter = 0
        Next listText

        ' Save our document into PDF format.
        dc.Save(documentPath, New PdfSaveOptions() With {.Compliance = PdfCompliance.PDF_A1a})

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
    End Sub
End Module