Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports System.Drawing

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			SeparateDocumentToImagePages()
		End Sub
		''' <summary>
		''' Load a document and save all pages as separate PNG &amp; Jpeg images.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/pagination-save-document-pages-as-png-jpg-jpeg-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub SeparateDocumentToImagePages()
			Dim filePath As String = "..\..\..\example.docx"
			Dim dc As DocumentCore = DocumentCore.Load(filePath)
			Dim folderPath As String = Path.GetFullPath("Result-files")
			Dim dp As DocumentPaginator = dc.GetPaginator()
			For i As Integer = 0 To dp.Pages.Count - 1
				Dim page As DocumentPage = dp.Pages(i)
				Directory.CreateDirectory(folderPath)

				' Save the each page as Bitmap.
				Dim bmp As Bitmap = page.Rasterize(300, SautinSoft.Document.Color.White)
				' Save the bitmap to PNG and JPEG.
				bmp.Save(folderPath & "\Page (PNG) - " & (i + 1).ToString() & ".png")
				bmp.Save(folderPath & "\Page (Jpeg) - " & (i + 1).ToString() & ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg)
			Next i
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(folderPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
