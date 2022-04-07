Imports System.IO
Imports SautinSoft.Document

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			SeparateDocumentToHtmlPages()
		End Sub
		''' <summary>
		''' Load a document and save all pages as separate HTML documents.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/pagination-save-document-pages-as-html-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub SeparateDocumentToHtmlPages()
			Dim filePath As String = "..\..\..\example.docx"
			Dim dc As DocumentCore = DocumentCore.Load(filePath)
			Dim folderPath As String = Path.GetFullPath("Result-files")
			Dim dp As DocumentPaginator = dc.GetPaginator()
			For i As Integer = 0 To dp.Pages.Count - 1
				Dim page As DocumentPage = dp.Pages(i)
				Directory.CreateDirectory(folderPath)

				' Save the each page into HTML format.
				page.Save(folderPath & "\Page - " & (i + 1).ToString() & ".html", SaveOptions.HtmlFixedDefault)
			Next i
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(folderPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
