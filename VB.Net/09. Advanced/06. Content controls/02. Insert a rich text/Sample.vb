Imports System
Imports System.Text
Imports SautinSoft.Document
Imports SautinSoft.Document.CustomMarkups
Imports SautinSoft.Document.Drawing
Imports System.IO

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			InsertRichText()
		End Sub
		''' <summary>
		''' Inserting a Rich text content control.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/content-controls-insert-rich-text-net-csharp-vb.php
		''' </remarks>

		Private Shared Sub InsertRichText()
			' Let's create a simple document.
			Dim dc As New DocumentCore()

			' Create a rich text content control.
			Dim rt As New BlockContentControl(dc, ContentControlType.RichText)
			dc.Sections.Add(New Section(dc, rt))

			' Add the content control properties.
			rt.Properties.Title = "Rich text content control."
			rt.Properties.Color = Color.Blue
			rt.Properties.LockDeleting = True
			rt.Properties.LockEditing = True
			rt.Document.DefaultCharacterFormat.FontColor = Color.Orange

			' Add new paragraphs with formatted text.
			rt.Blocks.Add(New Paragraph(dc, "Line 1"))
			rt.Blocks.Add(New Paragraph(dc, "Line 2"))
			rt.Blocks.Add(New Paragraph(dc, "Line 3"))

			' Add a picture and shape to the block.
			Dim pictPath As String = "..\..\..\banner_sautinsoft.jpg"
			Dim pict As New Picture(dc, InlineLayout.Inline(New Size(400, 100)), pictPath)
			Dim shp As New Shape(dc, Layout.Inline(New Size(4, 1, LengthUnit.Centimeter)))
			Dim im As New Paragraph(dc)
			rt.Blocks.Add(im)
			im.Inlines.Add(pict)
			Dim sh As New Paragraph(dc)
			rt.Blocks.Add(sh)
			sh.Inlines.Add(shp)

			rt.Blocks.Add(New Paragraph(dc, "Line 4"))

			' Save our document into DOCX format.
			Dim resultPath As String = "result.docx"
			dc.Save(resultPath, New DocxSaveOptions())

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace