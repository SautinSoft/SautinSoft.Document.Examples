Imports SautinSoft.Document
Imports System.Linq

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			InsertText()
		End Sub
		''' <summary>
		''' Insert a text into an existing PDF document in a specific position.
		''' </summary>
		''' </remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-text-to-pdf-document-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub InsertText()
			Dim filePath As String = "..\..\..\example.pdf"
			Dim fileResult As String = "Result.pdf"
			Dim dc As DocumentCore = DocumentCore.Load(filePath)

			' Find a position to insert text. Before this text: "> in this position".
			 Dim cr As ContentRange = dc.Content.Find("> in this position").FirstOrDefault()


			' Insert new text.
			If cr IsNot Nothing Then
				cr.Start.Insert("New text!")
			End If
			dc.Save(fileResult)
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(fileResult) With {.UseShellExecute = True})

		End Sub
	End Class
End Namespace
