Imports System
Imports System.IO
Imports SautinSoft.Document

Namespace Example
    Friend Class Program
        Shared Sub Main(ByVal args() As String)
            DifferentHeaderAndFooters()
        End Sub
        ''' <summary>
        ''' Creates a document with different headers: on first page, default and in another section.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-docx-document-with-different-headers-net-csharp-vb.php
        ''' </remarks>
        Private Shared Sub DifferentHeaderAndFooters()
            Dim documentPath As String = "DifferentHeaders.docx"

            Dim dc As New DocumentCore()
            Dim section1 As New Section(dc)
            dc.Sections.Add(section1)

            ' Create headers.
            Dim defaultHeader As HeaderFooter = CreateDefaultHeader(dc)
            Dim firstHeader As HeaderFooter = CreateFirstHeader(dc)

            ' Add the headers into the collection.
            section1.HeadersFooters.Add(defaultHeader)
            section1.HeadersFooters.Add(firstHeader)

            ' Set that main section has a different header on the first page.
            section1.PageSetup.TitlePage = True

            ' Add some content.
            Dim p As New Paragraph(dc, "My father's family name being Pirrip, and my Christian name Philip, " & "my infant tongue could make of both names nothing longer or more explicit than Pip. So, I called myself " & "Pip, and came to be called Pip. I give Pirrip as my father's family name, on the authority of his " & "tombstone and my sister, -Mrs. Joe Gargery, who married the blacksmith.")
            Dim totalPars As Integer = 15

            For i As Integer = 0 To totalPars - 1
                section1.Blocks.Add(p.Clone(True))
            Next i

            ' Add another header and content in the document.

            ' Let's create a one more section.
            Dim section2 As New Section(dc)
            dc.Sections.Add(section2)

            ' Create a new header.
            Dim header As New HeaderFooter(dc, HeaderFooterType.HeaderDefault)
            header.Content.Start.Insert("Chapter II - another header.", New CharacterFormat() With {
                .FontColor = Color.Blue,
                .Size = 18
            })

            ' Add the header into the Section2.
            section2.HeadersFooters.Add(header)

            ' Add some content.
            section2.Content.Start.Insert("My sister, Mrs. Joe Gargery, was more than twenty years older than I, " & "and had established a great reputation with herself and the neighbors because she had brought me " & "up ""by hand."" Having at that time to find out for myself what the expression meant, and knowing " & "her to have a hard and heavy hand, and to be much in the habit of laying it upon her husband as well " & "as upon me, I supposed that Joe Gargery and I were both brought up by hand.")

            ' Save the document into DOCX format.
            dc.Save(documentPath)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
        End Sub
        Private Shared Function CreateDefaultHeader(ByVal dc As DocumentCore) As HeaderFooter
            Dim header As New HeaderFooter(dc, HeaderFooterType.HeaderDefault)
            header.Content.Start.Insert("Chapter I - This is the default header", New CharacterFormat() With {
                .FontName = "Verdana",
                .Size = 18.0,
                .FontColor = Color.Orange
            })
            Return header
        End Function

        Private Shared Function CreateFirstHeader(ByVal dc As DocumentCore) As HeaderFooter
            Dim header As New HeaderFooter(dc, HeaderFooterType.HeaderFirst)
            header.Content.Start.Insert("Charles Dickens. Great Expectations - 1st page header", New CharacterFormat() With {
                .FontName = "Verdana",
                .Size = 18.0,
                .FontColor = Color.DarkGreen
            })
            Return header
        End Function
    End Class
End Namespace
