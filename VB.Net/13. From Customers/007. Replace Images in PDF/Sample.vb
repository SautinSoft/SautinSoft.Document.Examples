Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports SautinSoft.Document.Tables
Imports System.IO
Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			ReplaceImagesInPdf()
		End Sub

		''' <summary>
		''' How to replace images in PDF document.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-replace-images-in-pdf-in-csharp-vb-net.php
		''' </remarks>
		Private Shared Sub ReplaceImagesInPdf()

			' Path to a loadable document.
			Dim loadPath As String = "..\..\..\example.pdf"
			Dim pictPath As String = "..\..\..\replaceNA.jpg"

			' Load a document intoDocumentCore.
			Dim dc As DocumentCore = DocumentCore.Load(loadPath)

			' Load the Picture from a file.
			Dim picture As New Picture(dc, InlineLayout.Inline(New Size()), pictPath)

			' Find all pictures in the document.
			For Each el As Element In dc.GetChildElements(True, ElementType.Picture).Reverse()
				If TypeOf el Is Picture Then
					' Сopy all properties of the found picture and assign these properties to the new picture.
					' If you do not do this, the picture may be inserted into an arbitrary place in the document. 
					If TypeOf (CType(el, Picture)).Layout Is FloatingLayout Then
						Dim old As FloatingLayout = CType(CType(el, Picture).Layout, FloatingLayout)
						picture.Layout = FloatingLayout.Floating(old.HorizontalPosition, old.VerticalPosition, old.Size)
					End If

					' Replace picture.
					el.Content.Replace(picture.Content)
				End If
			Next el

			' Save our document into PDF format.
			Dim savePath As String = "replaced.pdf"
			dc.Save(savePath)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(loadPath) With {.UseShellExecute = True})
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(savePath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
