Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports System.Linq
Imports System.Text.RegularExpressions

Namespace Sample
	Friend Class Sample

		Shared Sub Main(ByVal args() As String)
			FindTextAndReplaceImage()
		End Sub


        ''' <summary>
        ''' Find Text and replace it with a Picture using ContentRange.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/find-text-replace-image-content-net-csharp-vb.php
        ''' </remarks>
        Public Shared Sub FindTextAndReplaceImage()
			' Path to a loadable document.
			Dim loadPath As String = "..\..\..\Critique_signature.docx"
			Dim pictPath As String = "..\..\..\Smile.png"

			' Load a document intoDocumentCore.
			Dim dc As DocumentCore = DocumentCore.Load(loadPath)

			'Find "<signature>" Text and Replace everywhere with the "Smile.png"
			' Please note, Reverse() makes sure that action replace not affects to Find().
			Dim regex As New Regex("<signature>", RegexOptions.IgnoreCase)
			Dim picture As New Picture(dc, InlineLayout.Inline(New Size(50, 50)), pictPath)
			For Each item As ContentRange In dc.Content.Find(regex).Reverse()
				item.Replace(picture.Content)
			Next item

			' Save our document into PDF format.
			Dim savePath As String = "..\..\..\Replaced_signature.pdf"
			dc.Save(savePath, New PdfSaveOptions())

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(loadPath) With {.UseShellExecute = True})
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(savePath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
