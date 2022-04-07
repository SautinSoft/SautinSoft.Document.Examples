Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Here we'll show two ways to create PDF document from a scratch.
			' Use any of them, which is more clear to you.

			' 1. With help of DocumentBuilder (wizard).
			CreatePdfUsingDocumentBuilder()

			' 2. With Document Object Model (DOM) directly.
			CreatePdfUsingDOM()
		End Sub
		''' <summary>
		''' Creates a new PDF document using DocumentBuilder wizard.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-pdf-document-net-csharp-vb.php
		''' </remarks>
		Public Shared Sub CreatePdfUsingDocumentBuilder()
			' Set a path to our document.
			Dim docPath As String = "Result-DocumentBuilder.pdf"

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
			db.Write("This is a first line in 1st paragraph!")
			' Add a line break into the 1st paragraph.
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)
			' Add 2nd line to the 1st paragraph, create 2nd paragraph.
			db.Writeln("Let's type a second line.")
			' Specify the paragraph alignment.
			TryCast(section.Blocks(0), Paragraph).ParagraphFormat.Alignment = HorizontalAlignment.Center

			' Add text into the 2nd paragraph.
			db.CharacterFormat.ClearFormatting()
			db.CharacterFormat.Size = 25
			db.CharacterFormat.FontColor = Color.Blue
			db.CharacterFormat.Bold = True
			db.Write("This is a first line in 2nd paragraph.")
			' Insert a line break into the 2nd paragraph.
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)
			' Insert 2nd line with own formatting to the 2nd paragraph.
			db.CharacterFormat.Size = 20
			db.CharacterFormat.FontColor = Color.DarkGreen
			db.CharacterFormat.UnderlineStyle = UnderlineType.Single
			db.CharacterFormat.Bold = False
			db.Write("This is a second line.")

			' Add a graphics figure into the paragraph.
			db.CharacterFormat.ClearFormatting()
			Dim shape As Shape = db.InsertShape(SautinSoft.Document.Drawing.Figure.SmileyFace, New SautinSoft.Document.Drawing.Size(50, 50, LengthUnit.Millimeter))
			' Specify outline and fill.
			shape.Outline.Fill.SetSolid(New SautinSoft.Document.Color("#358CCB"))
			shape.Outline.Width = 3
			shape.Fill.SetSolid(SautinSoft.Document.Color.Orange)

			' Save the document to the file in PDF format.
			dc.Save(docPath, New PdfSaveOptions() With {.Compliance = PdfCompliance.PDF_A1a})

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(docPath) With {.UseShellExecute = True})
		End Sub


		''' <summary>
		''' Creates a new PDF document using DOM directly.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-pdf-document-net-csharp-vb.php
		''' </remarks>
		Public Shared Sub CreatePdfUsingDOM()
			' Set a path to our document.
			Dim docPath As String = "Result-DocumentCore.pdf"

			' Create a new document.
			Dim dc As New DocumentCore()

			' Add new section.
			Dim section As New Section(dc)
			dc.Sections.Add(section)

			' Let's set page size A4.
			section.PageSetup.PaperType = PaperType.A4

			' Add two paragraphs            
			Dim par1 As New Paragraph(dc)
			par1.ParagraphFormat.Alignment = HorizontalAlignment.Center
			section.Blocks.Add(par1)

			' Let's create a characterformat for text in the 1st paragraph.
			Dim cf As New CharacterFormat() With {
				.FontName = "Verdana",
				.Size = 16,
				.FontColor = Color.Orange
			}
			Dim run1 As New Run(dc, "This is a first line in 1st paragraph!")
			run1.CharacterFormat = cf
			par1.Inlines.Add(run1)

			' Let's add a line break into the 1st paragraph.
			par1.Inlines.Add(New SpecialCharacter(dc, SpecialCharacterType.LineBreak))
			' Copy the formatting.
			Dim run2 As Run = run1.Clone()
			run2.Text = "Let's type a second line."
			par1.Inlines.Add(run2)

			' Add 2nd paragraph.
			Dim par2 As Paragraph = New Paragraph(dc, New Run(dc, "This is a first line in 2nd paragraph.", New CharacterFormat() With {
				.Size = 25,
				.FontColor = Color.Blue,
				.Bold = True
			}))
			section.Blocks.Add(par2)
			Dim lBr As New SpecialCharacter(dc, SpecialCharacterType.LineBreak)
			par2.Inlines.Add(lBr)
			Dim run3 As Run = New Run(dc, "This is a second line.", New CharacterFormat() With {
				.Size = 20,
				.FontColor = Color.DarkGreen,
				.UnderlineStyle = UnderlineType.Single
			})
			par2.Inlines.Add(run3)

			' Add a graphics figure into the paragraph.
			Dim shape As New Shape(dc, New InlineLayout(New SautinSoft.Document.Drawing.Size(50, 50, LengthUnit.Millimeter)))
			' Specify outline and fill.
			shape.Outline.Fill.SetSolid(New SautinSoft.Document.Color("#358CCB"))
			shape.Outline.Width = 3
			shape.Fill.SetSolid(SautinSoft.Document.Color.Orange)
			shape.Geometry.SetPreset(Figure.SmileyFace)
			par2.Inlines.Add(shape)

			' Save the document to the file in PDF format.
			dc.Save(docPath, New PdfSaveOptions() With {.Compliance = PdfCompliance.PDF_A1a})

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(docPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
