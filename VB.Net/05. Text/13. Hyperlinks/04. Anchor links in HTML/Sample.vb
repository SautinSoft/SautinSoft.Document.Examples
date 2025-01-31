Imports System.Text
Imports System.Linq
Imports SautinSoft.Document

Namespace Sample
	Friend Class Sample
		Shared Sub Main(ByVal args() As String)
			' Get your free trial key here:   
			' https://sautinsoft.com/start-for-free/

			AnchorLinks()
		End Sub

		''' <summary>
		''' Insert anchor links inside the HTML page using C# and .NET
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-anchor-links-inside-an-html-page-using-csharp-vb-net.php
		''' </remarks>		
		Public Shared Sub AnchorLinks()
			' P.S. If you are using MS Word, to display bookmarks:
			' File -> Options -> Advanced -> On the "Show document content" check "Show bookmarks".
			Dim documentPath As String = "AnchorLinks.html"

			' Let's create a new document.
			Dim dc As New DocumentCore()
			dc.Sections.Add(New Section(dc))

			' Add an anchor links to the end of the document.
			dc.Sections(0).Blocks.Add(New Paragraph(dc, New Hyperlink(dc, "IdEnd", "Document End") With {.IsBookmarkLink = True}))

			' Add 100 paragraphs
			For i As Integer = 0 To 99
				dc.Sections(0).Blocks.Add(New Paragraph(dc, New Run(dc, $"Paragraph {i + 1}")))
			Next i

			dc.Sections(0).Blocks.Add(New Paragraph(dc, New BookmarkStart(dc, "IdEnd"), New Run(dc, "The document end."), New BookmarkEnd(dc, "IdEnd")))

			' Let's save the document as HTML.
			dc.Save(documentPath, New HtmlFlowingSaveOptions())

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
		End Sub

	End Class
End Namespace
