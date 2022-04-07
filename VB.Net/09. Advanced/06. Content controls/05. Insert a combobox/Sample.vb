Imports System
Imports System.Text
Imports SautinSoft.Document
Imports SautinSoft.Document.CustomMarkups
Imports SautinSoft.Document.Drawing
Imports System.IO

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			InsertCombobox()
		End Sub
		''' <summary>
		''' Inserting a Combo Box content control.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/content-controls-insert-combobox-net-csharp-vb.php
		''' </remarks>

		Private Shared Sub InsertCombobox()
			' Let's create a simple document.
			Dim dc As New DocumentCore()

			' Create a Combo Box content control.
			Dim combobox As New InlineContentControl(dc, ContentControlType.ComboBox)
			dc.Sections.Add(New Section(dc, New Paragraph(dc, New Run(dc, "Combo Box "), combobox)))

			' Set common the content control properties.
			combobox.Properties.Title = "Combo Box"
			combobox.Properties.LockDeleting = True
			combobox.Properties.CharacterFormat.FontColor = Color.Blue

			' Add combox's list items.
			combobox.Properties.ListItems.Add(New ContentControlListItem("One", "1"))
			combobox.Properties.ListItems.Add(New ContentControlListItem("Two", "2"))
			combobox.Properties.ListItems.Add(New ContentControlListItem("Three", "3"))
			combobox.Properties.ListItems.Add(New ContentControlListItem("Four", "4"))

			' Set a default selected item.
			combobox.Properties.SelectedListItem = combobox.Properties.ListItems(2)

			' Save our document into DOCX format.
			Dim resultPath As String = "result.docx"
			dc.Save(resultPath, New DocxSaveOptions())

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace