Imports System
Imports SautinSoft.Document
Imports System.Text

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			InsertingBookmark()
		End Sub
		''' <summary>
		''' How to insert a Bookmark in a document using DocumentBuilder.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-bookmark.php
		''' </remarks>

		Private Shared Sub InsertingBookmark()
			Dim dc As New DocumentCore()
			Dim db As New DocumentBuilder(dc)

			Dim resultPath As String = "result.docx"

			' Insert the formatted text into the document using DocumentBuilder.
			db.CharacterFormat.FontName = "Verdana"
			db.CharacterFormat.Size = 16
			db.CharacterFormat.FontColor = Color.Orange
			db.Writeln("This text is inserted by the DocumentBuilder.Write method with formatting.")

			' Marks the current position in the document as a 1st bookmark start.
			db.StartBookmark("Firstbookmark")
			db.CharacterFormat.Italic = True
			db.CharacterFormat.Size = 12
			db.CharacterFormat.FontColor = Color.Blue
			db.Writeln("The text inside the bookmark 'Firstbookmark' is inserted by the DocumentBuilder.Writeln method.")

			' Marks the current position in the document as a 1st bookmark end.
			db.EndBookmark("Firstbookmark")

			' Insert text after the 1st bookmark.
			db.CharacterFormat.Italic = False
			db.CharacterFormat.Size = 16
			db.CharacterFormat.FontColor = Color.Orange
			db.Writeln("DocumentBuilder.EndBookmark method with the same name points to the end of the bookmark.")

			' Marks the current position in the document as a 2nd bookmark start.
			db.StartBookmark("Secondbookmark")
			db.CharacterFormat.Italic = True
			db.CharacterFormat.Size = 12
			db.CharacterFormat.FontColor = Color.Blue
			db.Writeln("Incorrectly spelled bookmarks or bookmarks with duplicate names will be ignored when saving the document.")

			' Marks the current position in the document as a 2nd bookmark end.
			db.EndBookmark("Secondbookmark")

			' Save our document into DOCX format.
			dc.Save(resultPath, New DocxSaveOptions())

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace