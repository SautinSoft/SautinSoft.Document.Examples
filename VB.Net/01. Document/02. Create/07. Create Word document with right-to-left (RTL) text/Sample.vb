Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free trial key here:   
			' https://sautinsoft.com/start-for-free/

			' "Right to left" (RTL) refers to the direction in which text is written and read, starting from the right side of the page or line and moving to the left.
			' This writing direction is used in languages such as Arabic, Hebrew, Persian, and Urdu.
			' It also influences the layout of user interfaces and documents in these languages

			Create_WORD_RTL()

		End Sub
		''' <summary>
		''' Right-to-Left. Converting Word file to PDF without losing any formatting.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-RTL-word-document-net-csharp-vb.php
		''' </remarks>
		Public Shared Sub Create_WORD_RTL()
			' Set a path to our document.
			Dim docPath As String = "Right-to-Left.docx"

			' Create a new document and DocumentBuilder.
			Dim dc As New DocumentCore()
			Dim db As New DocumentBuilder(dc)

			' Set page size A4.
			Dim section As Section = db.Document.Sections(0)
			section.PageSetup.PaperType = PaperType.A4

			' Add 1st paragraph with formatted text.
			db.CharacterFormat.FontName = "Verdana"
			db.CharacterFormat.Size = 16
			db.CharacterFormat.FontColor = Color.Orange
			db.Write("أخذ عن موالية الإمتعاض")
			' Add a line break into the 1st paragraph.
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)
			' Add 2nd line to the 1st paragraph, create 2nd paragraph.
			db.Writeln("של תיבת תרומה מלא")
			' Specify the paragraph alignment.
			TryCast(section.Blocks(0), Paragraph).ParagraphFormat.Alignment = HorizontalAlignment.Center

			' Add text into the 2nd paragraph.
			db.CharacterFormat.ClearFormatting()
			db.CharacterFormat.Size = 25
			db.CharacterFormat.FontColor = Color.Blue
			db.CharacterFormat.Bold = True
			db.Write("اُردُو حُرُوفِ تَہَجِّی‌")
			' Insert a line break into the 2nd paragraph.
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)
			' Insert 2nd line with own formatting to the 2nd paragraph.
			db.CharacterFormat.Size = 20
			db.CharacterFormat.FontColor = Color.DarkGreen
			db.CharacterFormat.UnderlineStyle = UnderlineType.Single
			db.CharacterFormat.Bold = False
			db.Write("هلو می فریند. ")

			' Add a graphics figure into the paragraph.
			db.CharacterFormat.ClearFormatting()
			Dim shape As Shape = db.InsertShape(SautinSoft.Document.Drawing.Figure.SmileyFace, New SautinSoft.Document.Drawing.Size(50, 50, LengthUnit.Millimeter))
			' Specify outline and fill.
			shape.Outline.Fill.SetSolid(New SautinSoft.Document.Color(53, 140, 203))
			shape.Outline.Width = 3
			shape.Fill.SetSolid(SautinSoft.Document.Color.White)

			' Save the document to the file in DOCX format.
			dc.Save(docPath, New DocxSaveOptions())

			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(docPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
