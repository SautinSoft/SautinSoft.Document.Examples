Imports System
Imports System.Text
Imports SautinSoft.Document
Imports SautinSoft.Document.CustomMarkups
Imports SautinSoft.Document.Drawing
Imports System.IO

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			InsertCheckBox()
		End Sub
		''' <summary>
		''' Inserting a Check box content control.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/content-controls-insert-checkbox-net-csharp-vb.php
		''' </remarks>

		Private Shared Sub InsertCheckBox()
			Dim dc As New DocumentCore()

			Dim checkbox As New InlineContentControl(dc, ContentControlType.CheckBox)

			' Set the checkbox properties.
			checkbox.Properties.Title = "Click me"
			checkbox.Properties.Checked = True
			checkbox.Properties.LockDeleting = True
			checkbox.Properties.CharacterFormat.Size = 24

			' Override default checkbox appearance.
			checkbox.Properties.CheckedSymbol.FontName = "Courier New"
			checkbox.Properties.CheckedSymbol.Character = "X"c
			checkbox.Properties.UncheckedSymbol.FontName = "Courier New"
			checkbox.Properties.UncheckedSymbol.Character = "O"c

			dc.Sections.Add(New Section(dc, New Paragraph(dc, New Run(dc, "Click me => ", New CharacterFormat() With {
				.Size = 24,
				.FontColor = New Color("#3399FF")
			}), checkbox)))

			' Save our document into DOCX format.
			Dim resultPath As String = "result.docx"
			dc.Save(resultPath, New DocxSaveOptions())

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace