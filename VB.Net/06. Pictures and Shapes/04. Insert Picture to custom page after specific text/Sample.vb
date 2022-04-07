Imports System.Linq
Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			InsertPictureToCustomPageAndPosition()
		End Sub
		''' <summary>
		''' Insert a picture to custom page and position into existing DOCX document.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/insert-picture-jpg-image-to-custom-docx-page-and-position-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub InsertPictureToCustomPageAndPosition()
			' In this example we'll insert the picture to 1st after the word "Sign:".            

			Dim inpFile As String = "..\..\..\example.docx"
			Dim outFile As String = "Result.docx"
			Dim pictFile As String = "..\..\..\picture.jpg"

			Dim dc As DocumentCore = DocumentCore.Load(inpFile)
			Dim dp As DocumentPaginator = dc.GetPaginator()

			' Find the text "Sign:" on the 1st page.
			Dim cr As ContentRange = dp.Pages(0).Content.Find("Sign:").LastOrDefault()
			If cr IsNot Nothing Then
				Dim pic As New Picture(dc, pictFile)
				cr.End.Insert(pic.Content)
			End If
			' Save the document as new DOCX and open it.
			dc.Save(outFile)
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
