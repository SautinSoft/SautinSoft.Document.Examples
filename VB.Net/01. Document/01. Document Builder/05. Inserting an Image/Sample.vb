Imports System
Imports SautinSoft.Document
Imports System.Text
Imports SautinSoft.Document.Drawing

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			InsertingImage()
		End Sub
		''' <summary>
		''' Insert an image and shape inline or in the specified position using DocumentBuilder.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-image.php
		''' </remarks>

		Private Shared Sub InsertingImage()
			Dim dc As New DocumentCore()
			Dim db As New DocumentBuilder(dc)

			Dim resultPath As String = "result.docx"
			Dim pictPath As String = "..\..\..\logo.png"

			' Insert the formatted text into the document using DocumentBuilder.
			db.CharacterFormat.FontName = "Courier"
			db.CharacterFormat.Size = 17.0F
			db.CharacterFormat.Italic = True
			db.CharacterFormat.FontColor = Color.Orange
			db.Writeln("Insert an Image and Shape using DocumentBuilder.")

			' Images:
			' 1st way: Insert an Inline image into the document.
			' Specify the image size and rotation (if required).
			Dim pict1 As Picture = db.InsertImage(pictPath, New Size(100, 30, LengthUnit.Millimeter))
			pict1.Rotation = -3

			' 2nd way: Insert a Floating image from a file at the specified position.
			Dim pict2 As Picture = db.InsertImage(pictPath, New HorizontalPosition(1, LengthUnit.Centimeter, HorizontalPositionAnchor.LeftMargin), New VerticalPosition(8, LengthUnit.Centimeter, VerticalPositionAnchor.TopMargin), WrappingStyle.InFrontOfText)

			' Shapes:
			' 1st way: Insert an Inline shape.
			Dim shp1 As Shape = db.InsertShape(Figure.SmileyFace, New Size(3, 3, LengthUnit.Centimeter))

			' 2nd way: Insert a Floating shape with specified position, size and text wrap style.
			Dim size1 As New Size(7, 6, LengthUnit.Centimeter)
			Dim shp2 As Shape = db.InsertShape(Figure.RoundRectangle, New HorizontalPosition(8, LengthUnit.Centimeter, HorizontalPositionAnchor.LeftMargin), New VerticalPosition(10, LengthUnit.Centimeter, VerticalPositionAnchor.TopMargin), WrappingStyle.InFrontOfText, New Size(7, 6, LengthUnit.Centimeter))

			shp2.Fill.SetSolid(Color.White)

			' Move the "cursor" position inside the shape content.
			db.MoveTo(shp2.Text.Blocks.Content.Start)
			db.CharacterFormat.FontColor = Color.Green
			db.CharacterFormat.Size = 26.0F
			db.Writeln("Text inside Shape.")

			' Move the "cursor" back to the paragraph with "shp1".
			db.MoveTo(shp1.Content.End)

			' Save our document into DOCX format.
			dc.Save(resultPath, New DocxSaveOptions())

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace