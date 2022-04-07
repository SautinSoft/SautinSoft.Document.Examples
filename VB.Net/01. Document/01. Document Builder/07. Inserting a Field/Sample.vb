Imports System
Imports SautinSoft.Document
Imports System.Text
Imports SautinSoft.Document.Drawing

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			InsertingField()
		End Sub
		''' <summary>
		''' Generate document with forms and fields using DocumentBuilder.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-field.php
		''' </remarks>

		Private Shared Sub InsertingField()
			Dim dc As New DocumentCore()
			Dim db As New DocumentBuilder(dc)

			Dim resultPath As String = "result.pdf"
			Dim items() As String = {"One", "Two", "Three", "Four", "Five"}

			' Insert the formatted text into the document using DocumentBuilder.
			db.CharacterFormat.FontName = "Verdana"
			db.CharacterFormat.Size = 16
			db.CharacterFormat.FontColor = Color.Orange
			db.Writeln("Generate document with forms and fields using DocumentBuilder.")
			db.CharacterFormat.ClearFormatting()

			db.Write("{ TIME   \* MERGEFORMAT } - ")
			db.InsertField("TIME")
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)

			db.InsertTextInput("TextInput", FormTextType.RegularText, "", "Insert Text Input", 0)
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)

			db.InsertCheckBox("CheckBox", True, 0)
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)

			db.InsertComboBox("DropDown", items, 3)

			' Save our document into PDF format.
			dc.Save(resultPath, New PdfSaveOptions())

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace