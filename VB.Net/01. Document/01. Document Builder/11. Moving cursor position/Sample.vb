Imports System
Imports SautinSoft.Document
Imports System.Text
Imports SautinSoft.Document.Drawing
Imports SautinSoft.Document.Tables

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			MovingCursor()
		End Sub
		''' <summary>
		''' Moving the current cursor position in the document using DocumentBuilder.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-moving-cursor.php
		''' </remarks>

		Private Shared Sub MovingCursor()
			Dim dc As New DocumentCore()
			Dim db As New DocumentBuilder(dc)

			db.MoveToHeaderFooter(HeaderFooterType.HeaderDefault)
			db.Writeln("Moved the cursor to the header and inserted this text.")

			db.MoveToDocumentStart()
			db.CharacterFormat.Size = 16
			db.CharacterFormat.FontColor = Color.Blue
			db.Writeln("Moved the cursor to the start of the document.")

			' Marks the current position in the document as a 1st bookmark start.
			db.StartBookmark("Firstbookmark")
			db.CharacterFormat.Italic = True
			db.CharacterFormat.Size = 14
			db.CharacterFormat.FontColor = Color.Red
			db.Writeln("The text inside the 'Bookmark' is inserted by the DocumentBuilder.Writeln method.")
			' Marks the current position in the document as a 1st bookmark end.
			db.EndBookmark("Firstbookmark")

			db.MoveToBookmark("Firstbookmark", True, False)
			db.Writeln("Moved the cursor to the start of the Bookmark.")

			db.CharacterFormat.FontColor = Color.Black
			Dim f1 As Field = db.InsertField("DATE")
			db.MoveToField(f1, False)
			db.Write("Before the field")

			' Moving to the Header and insert the table with three cells.
			db.MoveToHeaderFooter(HeaderFooterType.HeaderDefault)
			db.StartTable()
			db.TableFormat.PreferredWidth = New TableWidth(LengthUnitConverter.Convert(6, LengthUnit.Inch, LengthUnit.Point), TableWidthUnit.Point)
			db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Green, 1)
			db.RowFormat.Height = New TableRowHeight(40, HeightRule.Exact)
			db.CharacterFormat.FontColor = Color.Green
			db.CharacterFormat.Italic = False
			db.InsertCell()
			db.Write("This is Row 1 Cell 1")
			db.InsertCell()
			db.Write("This is Row 1 Cell 2")
			db.InsertCell()
			db.Write("This is Row 1 Cell 3")
			db.EndTable()

			' Insert the text in the second cell in the sixth position.
			db.MoveToCell(0, 0, 1, 5)
			db.CharacterFormat.Size = 18
			db.CharacterFormat.FontColor = Color.Orange
			db.Write("InsertToCell")

			db.MoveToDocumentEnd()
			db.CharacterFormat.Size = 16
			db.CharacterFormat.FontColor = Color.Blue
			db.Writeln("Moved the cursor to the end of the document.")

			' Save our document into DOCX format.
			Dim resultPath As String = "result.docx"
			dc.Save(resultPath, New DocxSaveOptions())

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
