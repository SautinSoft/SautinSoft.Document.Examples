Imports System
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports System.IO
Imports System.Linq

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			WriteProtection()
		End Sub

		''' <summary>
		''' Create a write protected DOCX document.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/write-protection-options-net-csharp-vb.php
		''' </remarks>
		Public Shared Sub WriteProtection()
			Dim filePath As String = "ProtectedDocument.docx"

			Dim dc As New DocumentCore()

			' Insert paragraphs into the document.
			dc.Sections.Add(New Section(dc, New Paragraph(dc, "This document has been opened in read only mode."),
										New Paragraph(dc, "To keep your changes, you 'll need to save the document with a new name or in a different location."),
										New Paragraph(dc, "To make changes to the current document, restart with the password '12345'.")))

			' Sets the write protection password "12345".
			Dim protection As DocumentWriteProtection = dc.WriteProtection
			protection.SetPassword("12345")

			' Save a document as the DOCX file with write protection options.
			dc.Save(filePath)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
