Imports System
Imports SautinSoft.Document
Imports System.Text


Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			InsertingText()
		End Sub
		''' <summary>
		''' Create a document and insert a string of text using DocumentBuilder.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-text.php
		''' </remarks>

		Private Shared Sub InsertingText()
			Dim dc As New DocumentCore()
			Dim db As New DocumentBuilder(dc)

			Dim resultPath As String = "result.pdf"

			' Insert the formatted text into the document using DocumentBuilder.
			db.CharacterFormat.FontName = "Verdana"
			db.CharacterFormat.Size = 55.5F
			db.CharacterFormat.AllCaps = True
			db.CharacterFormat.Italic = True
			db.CharacterFormat.FontColor = Color.Orange
			db.Write("insert a text using")

			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)

			db.CharacterFormat.Size = 52.5F
			db.CharacterFormat.FontColor = Color.Blue
			db.CharacterFormat.AllCaps = False
			db.CharacterFormat.Italic = False
			db.Write("DocumentBuilder")

			' Save the document to the file in PDF format.
			dc.Save(resultPath, New PdfSaveOptions() With {.Compliance = PdfCompliance.PDF_A1a})

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace