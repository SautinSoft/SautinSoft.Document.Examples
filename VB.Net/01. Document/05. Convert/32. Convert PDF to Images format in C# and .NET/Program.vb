Imports SautinSoft.Document

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			ConvertPDFtoImages()
		End Sub

		''' <summary>
		''' Convert PDF to Images (file to file).
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-pdf-to-images-in-csharp-vb.php
		''' </remarks>
		Private Shared Sub ConvertPDFtoImages()
			' Path to a document where to extract pictures.
			' By the way: You may specify DOCX, HTML, RTF files also.
			Dim dc As DocumentCore = DocumentCore.Load("..\..\..\example.pdf")

			' PaginationOptions allow to know, how many pages we have in the document.
			Dim dp As DocumentPaginator = dc.GetPaginator(New PaginatorOptions())

			' Each document page will be saved in its own image format: PNG, JPEG, TIFF with different DPI.
			dp.Pages(0).Save("..\..\..\example.png", New ImageSaveOptions() With {
				.DpiX = 800,
				.DpiY = 800
			})
			dp.Pages(1).Save("..\..\..\example.jpeg", New ImageSaveOptions() With {
				.DpiX = 400,
				.DpiY = 400
			})
			dp.Pages(2).Save("..\..\..\example.tiff", New ImageSaveOptions() With {
				.DpiX = 650,
				.DpiY = 650
			})

		End Sub
	End Class
End Namespace
