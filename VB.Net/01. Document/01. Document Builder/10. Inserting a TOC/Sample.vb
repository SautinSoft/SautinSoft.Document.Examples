Imports System
Imports SautinSoft.Document
Imports System.Text

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			InsertingToc()
		End Sub
		''' <summary>
		''' Insert a TOC (Table of Contents) field into the document using DocumentBuilder.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-toc.php
		''' </remarks>

		Private Shared Sub InsertingToc()
			Dim dc As New DocumentCore()
			Dim db As New DocumentBuilder(dc)

			' Insert the formatted text into the document using DocumentBuilder.
			db.CharacterFormat.FontName = "Verdana"
			db.CharacterFormat.Size = 16
			db.Writeln("Table of Contents.")
			db.CharacterFormat.ClearFormatting()

			' Insert Table of Contents field into the document at the current position.
			Dim toe As TableOfEntries = db.InsertTableOfContents("\o ""1-3"" \h")
			' For information about switches, see the description on the page above.

			' Add the text and divide it into headings.
			db.InsertSpecialCharacter(SpecialCharacterType.PageBreak)
			Dim Heading1Style As ParagraphStyle = CType(Style.CreateStyle(StyleTemplateType.Heading1, dc), ParagraphStyle)
			dc.Styles.Add(Heading1Style)
			db.ParagraphFormat.Style = Heading1Style
			db.Writeln("Heading 1")
			db.ParagraphFormat.ClearFormatting()
			db.Writeln("Some text Heading 1 Some text Heading 1 Some text Heading 1 Some text Heading 1 Some text Heading 1 Some text Heading 1" & "Some text Heading 1 Some text Heading 1 Some text Heading 1 Some text Heading 1 Some text Heading 1 Some text Heading 1 ")

			Dim Heading2Style As ParagraphStyle = CType(Style.CreateStyle(StyleTemplateType.Heading2, dc), ParagraphStyle)
			dc.Styles.Add(Heading2Style)
			db.ParagraphFormat.Style = Heading2Style
			db.Writeln("Heading 1.1")
			db.ParagraphFormat.ClearFormatting()
			db.Writeln("Some text Heading 1.1 Some text Heading 1.1 Some text Heading 1.1 Some text Heading 1.1 Some text Heading 1.1" & " Some text Heading 1.1 Some text Heading 1.1 Some text Heading 1.1 Some text Heading 1.1")
			db.ParagraphFormat.Style = Heading2Style
			db.Writeln("Heading 1.2")
			db.ParagraphFormat.ClearFormatting()
			db.Writeln("Some text Heading 1.2 Some text Heading 1.2 Some text Heading 1.2 Some text Heading 1.2 Some text Heading 1.2" & " Some text Heading 1.2 Some text Heading 1.2 Some text Heading 1.2 ")

			Dim Heading3Style As ParagraphStyle = CType(Style.CreateStyle(StyleTemplateType.Heading3, dc), ParagraphStyle)
			dc.Styles.Add(Heading3Style)
			db.ParagraphFormat.Style = Heading3Style
			db.Writeln("Heading 1.1.1")
			db.ParagraphFormat.ClearFormatting()
			db.Writeln("Some text Heading 1.1.1 Some text Heading 1.1.1 Some text Heading 1.1.1 Some text Heading 1.1.1 " & " Some text Heading 1.1.1 Some text Heading 1.1.1 Some text Heading 1.1.1 ")
			db.ParagraphFormat.Style = Heading3Style
			db.Writeln("Heading 1.1.2")
			db.ParagraphFormat.ClearFormatting()
			db.Writeln("Some text 1.1.2 Some text 1.1.2 Some text 1.1.2 Some text 1.1.2")

			db.ParagraphFormat.Style = Heading1Style
			db.Writeln("Heading 2")
			db.ParagraphFormat.ClearFormatting()
			db.Writeln("Some text Heading 2 Some text Heading 2.")

			db.ParagraphFormat.Style = Heading1Style
			db.Writeln("Heading 3")
			db.ParagraphFormat.ClearFormatting()
			db.Writeln("Some text Heading 3 Some text Heading 3 Some text Heading 3 Some text Heading 3 Some text Heading 3" & "Some text Heading 3Some text Heading 3Some text Heading 3Some text Heading 3Some text Heading 3")

			db.ParagraphFormat.Style = Heading2Style
			db.Writeln("Heading 3.1")
			db.ParagraphFormat.ClearFormatting()
			db.Writeln("Some text Heading 3.1 Some text Heading 3.1 Some text Heading 3.1 Some text Heading 3.1 Some text Heading 3.1" & "Some text Heading 3.1 Some text Heading 3.1 Some text Heading 3.1")

			db.ParagraphFormat.Style = Heading2Style
			db.Writeln("Heading 3.2")
			db.ParagraphFormat.ClearFormatting()
			db.Writeln("Some text Heading 3.2 Some text Heading 3.2 Some text Heading 3.2 Some text Heading 3.2 Some text Heading 3.2")

			' Update the TOC field (table of contents).
			toe.Update()

			' Save our document into DOCX format.
			Dim resultPath As String = "result.docx"
			dc.Save(resultPath, New DocxSaveOptions())

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace