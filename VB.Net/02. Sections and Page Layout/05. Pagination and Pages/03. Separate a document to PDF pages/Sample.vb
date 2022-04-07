Imports System.IO
Imports SautinSoft.Document

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			SeparateDocumentToPdfPages()
		End Sub
		''' <summary>
		''' Load a document and save all pages as separate PDF documents.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/pagination-save-document-pages-as-pdf-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub SeparateDocumentToPdfPages()
			Dim filePath As String = "..\..\..\example.docx"
			Dim dc As DocumentCore = DocumentCore.Load(filePath)
			Dim folderPath As String = Path.GetFullPath("Result-files")
			Dim dp As DocumentPaginator = dc.GetPaginator()
			For i As Integer = 0 To dp.Pages.Count - 1
				Dim page As DocumentPage = dp.Pages(i)
				Directory.CreateDirectory(folderPath)

				' Save the each page into PDF format.
				page.Save(folderPath & "\Page - " & (i + 1).ToString() & ".pdf", SaveOptions.PdfDefault)
			Next i
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(folderPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
