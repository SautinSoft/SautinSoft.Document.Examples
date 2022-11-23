Imports SautinSoft.Document
Imports System

Namespace Sample
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			DeleteSpecifiedPageInDocument()
		End Sub
		''' <summary>
		''' How to delete the specified page in the document.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-delete-specified-page-in-document-net-csharp-vb.php
		''' </remarks>
		Public Shared Sub DeleteSpecifiedPageInDocument()
			Dim inpFile As String = "..\..\..\example.docx"
			Dim outFile As String = "result.docx"

			' Load a document into DocumentCore.
			Dim dc As DocumentCore = DocumentCore.Load(inpFile)

			' Divide the document into separate pages.
			Dim dp As DocumentPaginator = dc.GetPaginator()

			' Delete page number two.
			dp.Pages(1).Content.Delete()

			 ' Save our result as a DOCX file.
			dc.Save(outFile)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace