Imports SautinSoft.Document

Namespace Sample
    Friend Class Sample
        Shared Sub Main(ByVal args() As String)
            AddBookmarks()
        End Sub

        ''' <summary>
        ''' How add a text bounded by BookmarkStart and BookmarkEnd. 
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/bookmarks.php
        ''' </remarks>		
        Public Shared Sub AddBookmarks()
            ' P.S. If you are using MS Word, to display bookmarks:
            ' File -> Options -> Advanced -> On the "Show document content" check "Show bookmarks".
            Dim documentPath As String = "Bookmarks.docx"

            ' Let's create a new document.
            Dim dc As New DocumentCore()

            ' Add text bounded by BookmarkStart and BookmarkEnd.
            dc.Sections.Add(New Section(dc, New Paragraph(dc, New Run(dc, "Text before bookmark. "), New BookmarkStart(dc, "SimpleBookmark"), New Run(dc, "Text inside bookmark."), New BookmarkEnd(dc, "SimpleBookmark"), New Run(dc, " Text after bookmark.")), New Paragraph(dc, New SpecialCharacter(dc, SpecialCharacterType.PageBreak), New Hyperlink(dc, "SimpleBookmark", "Go to Simple Bookmark.") With {.IsBookmarkLink = True})))

            ' Modify text inside bookmark.
            dc.Bookmarks("SimpleBookmark").GetContent(False).Replace("Some text inside bookmark.")

            ' Let's save our document into DOCX format.
            dc.Save(documentPath)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
