Imports System
Imports SautinSoft.Document
Imports System.Text
Imports SautinSoft.Document.Drawing

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			InsertingHyperlink()
		End Sub
		''' <summary>
		''' Insert a hyperlink into a document using DocumentBuilder.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-hyperlink.php
		''' </remarks>

		Private Shared Sub InsertingHyperlink()
			Dim dc As New DocumentCore()
			Dim db As New DocumentBuilder(dc)

			Dim resultPath As String = "result.docx"

			' Insert the formatted text into the document.
			db.CharacterFormat.FontName = "Verdana"
			db.CharacterFormat.Size = 16
			db.Writeln("Insert a hyperlink into a document using DocumentBuilder.")

			' Inserts a Word field into a document.
			db.CharacterFormat.Size = 26
			db.CharacterFormat.FontColor = Color.Brown
			db.InsertField("DATE")
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)

			' Insert URL hyperlink.
			db.CharacterFormat.FontColor = Color.Blue
			db.CharacterFormat.UnderlineStyle = UnderlineType.Dashed
			db.InsertHyperlink("Welcome to SautinSoft!", "https://sautinsoft.com", False)

			db.InsertSpecialCharacter(SpecialCharacterType.PageBreak)

			' Insert a hyperlink inside a document as a bookmark.
			db.CharacterFormat.FontColor = Color.Brown
			db.CharacterFormat.UnderlineStyle = UnderlineType.DotDotDash
			db.InsertHyperlink("back to the field {DATE}", "DATE", True)

			' Save our document into DOCX format.
			dc.Save(resultPath, New DocxSaveOptions())

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
