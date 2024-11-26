Imports System.Linq
Imports SautinSoft.Document

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			ReplaceText()
		End Sub
		''' <summary>
		''' Replace a specific text in an existing DOCX document.
		''' </summary>
		''' </remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/replace-text-in-docx-document-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub ReplaceText()
			Dim filePath As String = "..\..\..\example.docx"
			Dim fileResult As String = "Result.docx"
			Dim searchText As String = "document"
			Dim replaceText As String = "book"
			Dim dc As DocumentCore = DocumentCore.Load(filePath)
			For Each cr As ContentRange In dc.Content.Find(searchText).Reverse()
				' Replace "document" to "book";
				' Mark "book" by yellow.
				cr.Replace(replaceText, New CharacterFormat() With {.BackgroundColor = Color.Yellow})
			Next cr
			dc.Save(fileResult)
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(fileResult) With {.UseShellExecute = True})

		End Sub
	End Class
End Namespace
