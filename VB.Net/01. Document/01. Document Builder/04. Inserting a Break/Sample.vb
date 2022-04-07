Imports System
Imports SautinSoft.Document
Imports System.Text


Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			InsertingBreak()
		End Sub
		''' <summary>
		''' Insert a Line Break, Column Break, Page Break using DocumentBuilder.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-break.php
		''' </remarks>

		Private Shared Sub InsertingBreak()
			Dim dc As New DocumentCore()
			Dim db As New DocumentBuilder(dc)

			Dim resultPath As String = "result.docx"
			db.PageSetup.TextColumns = New TextColumnCollection(2)

			' Insert the formatted text into the document using DocumentBuilder.
			db.CharacterFormat.FontName = "Verdana"
			db.CharacterFormat.Size = 16.5F
			db.CharacterFormat.AllCaps = True
			db.CharacterFormat.Italic = True
			db.CharacterFormat.FontColor = Color.Orange
			db.ParagraphFormat.LeftIndentation = 30
			db.Writeln("This paragraph has a Left Indentation of 30 points.")

			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)

			' Undo the previously applied formatting.
			db.ParagraphFormat.ClearFormatting()
			db.CharacterFormat.ClearFormatting()

			db.Writeln("After this paragraph insert a column break.")
			db.InsertSpecialCharacter(SpecialCharacterType.ColumnBreak)

			db.CharacterFormat.Italic = True
			db.CharacterFormat.FontColor = Color.DarkBlue
			db.CharacterFormat.Size = 20.0F
			db.Writeln("After this paragraph insert a page break.")
			db.InsertSpecialCharacter(SpecialCharacterType.PageBreak)

			' Save the document to the file in DOCX format.
			dc.Save(resultPath)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace