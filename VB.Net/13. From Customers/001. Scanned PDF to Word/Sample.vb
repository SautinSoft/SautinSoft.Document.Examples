Imports SautinSoft.Document

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			ScannedPdfToWord()
		End Sub

        ''' <summary>
        ''' The method converts a PDF document with scanned images to Word. But it works only if the PDF document contains a hidden text atop of the images.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-scanned-pdf-to-word-in-csharp-vb-net.php
        ''' </remarks>
        Private Shared Sub ScannedPdfToWord()
			' Actually there are a lot of PDF documents which looks like created using a scanner, 
			' but they also contain a hidden text atop of the contents. 
			' This hidden text duplicates the content of the scanned images. 
			' This is made specially to have the ability to perform the 'find' operation.

			' Our steps:
			' 1. Load the PDF with the these settings: 
			' - show hidden text;
			' - skip all images during the loading. 
			' 2. Change the font color to the 'Black' for the all text.
			' 3. Save the document as DOCX.
			Dim inpFile As String = "..\..\..\Scanned.pdf"
			Dim outFile As String = "Result.docx"

			Dim pdfLO As New PdfLoadOptions
			With pdfLO
				' 'Disabled' - Never load embedded fonts in PDF. Use the fonts with the same name installed at the system or similar by font metrics.
				' 'Enabled' - Always load embedded fonts in PDF.
				' 'Auto' - Load only embedded fonts missing in the system. In other case, use the system fonts.
				.PreserveEmbeddedFonts = PropertyState.Enabled
				.PreserveImages = False
				.ShowInvisibleText = True
			End With

			Dim dc As DocumentCore = DocumentCore.Load(inpFile, pdfLO)

            dc.DefaultCharacterFormat.FontColor = Color.Black
			For Each element As Element In dc.GetChildElements(True, ElementType.Paragraph)
				For Each inline As Inline In (TryCast(element, Paragraph)).Inlines
					If TypeOf inline Is Run Then
						TryCast(inline, Run).CharacterFormat.FontColor = Color.Black
					End If
				Next inline
				TryCast(element, Paragraph).CharacterFormatForParagraphMark.FontColor = Color.Black
			Next element
			dc.Save(outFile)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
