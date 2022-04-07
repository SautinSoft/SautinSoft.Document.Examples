Imports SautinSoft.Document

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' You can create the same document by using 3 ways:
			'
			'  + DocumentBuilder
			'  + DOM directly
			'  + DOM and ContentRange
			'
			' Choose any of them which you like.

			' Way 1:
			CreateUsingDocumentBuilder()

			' Way 2:
			CreateUsingDOM()

			' Way 3:
			CreateUsingContentRange()
		End Sub

		''' <summary>
		''' Creates a new document using DocumentBuilder and saves it in a desired format.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/create-document.php
		''' </remarks>
		Private Shared Sub CreateUsingDocumentBuilder()
			' Create a new document and DocumentBuilder.
			Dim dc As New DocumentCore()
			Dim db As New DocumentBuilder(dc)

			' Specify the formatting and insert text.
			db.CharacterFormat.FontName = "Verdana"
			db.CharacterFormat.Size = 65.5F
			db.CharacterFormat.FontColor = Color.Orange
			db.Write("Hello World!")

			' Save the document in DOCX format.
			Dim outFile As String = "DocumentBuilder.docx"
			dc.Save(outFile)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
		''' <summary>
		''' Creates a new document using DOM and saves it in a desired format.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/create-document.php
		''' </remarks>
		Private Shared Sub CreateUsingDOM()
			' Create a new document.
			Dim dc As New DocumentCore()

			' Create a new section,
			' add the section the document.
			Dim sect As New Section(dc)
			dc.Sections.Add(sect)

			' Create a new paragraph,
			' add the paragraph to the section.
			Dim par As New Paragraph(dc)
			sect.Blocks.Add(par)

			' Create a new run (text object),
			' add the run to the paragraph.
			Dim run As New Run(dc, "Hello World!")
			run.CharacterFormat.FontName = "Verdana"
			run.CharacterFormat.Size = 65.5F
			run.CharacterFormat.FontColor = Color.Orange
			par.Inlines.Add(run)

			' Save the document in PDF format.
			Dim outFile As String = "DOM.pdf"
			dc.Save(outFile)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub


		''' <summary>
		''' Creates a new document using DOM and ContentRange and saves it in a desired format.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/create-document.php
		''' </remarks>
		Private Shared Sub CreateUsingContentRange()
			' Create a new document.
			Dim dc As New DocumentCore()
			' Insert the formatted text into the document.
			dc.Content.End.Insert("Hello World!", New CharacterFormat() With {
				.FontName = "Verdana",
				.Size = 65.5F,
				.FontColor = Color.Orange
			})

			' Save the document in HTML format.
			Dim outFile As String = "ContentRange.html"
			dc.Save(outFile, New HtmlFixedSaveOptions() With {.Title = "ContentRange"})

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace