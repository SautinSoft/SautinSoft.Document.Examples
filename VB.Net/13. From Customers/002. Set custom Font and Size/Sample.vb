Imports System
Imports System.IO
Imports SautinSoft.Document

Namespace Sample
	Friend Class Sample
		Shared Sub Main(ByVal args() As String)
			SetCustomFontAndSize()
		End Sub
		''' <summary>
		''' Convert DOCX document to PDF file (set custom font, size and line spacing).
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-word-to-pdf-set-custom-font-and-size-csharp-vb-net.php
		''' </remarks>
		Public Shared Sub SetCustomFontAndSize()
			' Path to a loadable document.
			Dim inpFile As String = "..\..\..\example.docx"
			Dim outFile As String = "result set custom font.pdf"

			Dim dc As DocumentCore = DocumentCore.Load(inpFile)

			Dim singleFontName As String = "Times New Roman"
			Dim singleFontSize As Single = 8.0F
			Dim singleLineSpacing As Single = 0.8F

			dc.DefaultCharacterFormat.FontName = singleFontName
			dc.DefaultCharacterFormat.Size = singleFontSize

			For Each element As Element In dc.GetChildElements(True, ElementType.Run, ElementType.Paragraph)
				If TypeOf element Is Run Then
					TryCast(element, Run).CharacterFormat.FontName = singleFontName
					TryCast(element, Run).CharacterFormat.Size = singleFontSize
				ElseIf TypeOf element Is Paragraph Then
					TryCast(element, Paragraph).ParagraphFormat.LineSpacing = singleLineSpacing
				End If
			Next element
			dc.Save(outFile)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(inpFile) With {.UseShellExecute = True})
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
