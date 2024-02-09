Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Namespace Sample
	Friend Class Sample

		Shared Sub Main(ByVal args() As String)
			' Get your free 30 - day key here:   
			' https://sautinsoft.com/start-for-free/

			FullPageImage()
		End Sub

		''' <summary>
		''' How to add pictures into a document. 
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/how-to-map-image-csharp-net.php
		''' </remarks>
		Public Shared Sub FullPageImage()
			Dim documentPath As String = "pictures.docx"
			Dim pictPath As String = "..\..\..\image-png.png"

			' Let's create a simple document.
			Dim dc As New DocumentCore()

			' Add a new section, A5 Landscape, and custom page margins.
			Dim s As New Section(dc)
			s.PageSetup.PaperType = PaperType.A5
			s.PageSetup.Orientation = Orientation.Landscape
			s.PageSetup.PageMargins = New PageMargins() With {
				.Top = LengthUnitConverter.Convert(0, LengthUnit.Millimeter, LengthUnit.Point),
				.Right = LengthUnitConverter.Convert(0, LengthUnit.Inch, LengthUnit.Point),
				.Bottom = LengthUnitConverter.Convert(0, LengthUnit.Millimeter, LengthUnit.Point),
				.Left = LengthUnitConverter.Convert(0, LengthUnit.Inch, LengthUnit.Point)
			}
			dc.Sections.Add(s)

			' Create a new paragraph with picture.
			Dim par As New Paragraph(dc)
			s.Blocks.Add(par)

			' Our picture has InlineLayout - it doesn't have positioning by coordinates
			' and located as flowing content together with text (Run and other Inline elements).
			Dim pict1 As New Picture(dc, InlineLayout.Inline(New Size(s.PageSetup.PageWidth, s.PageSetup.PageHeight)), pictPath)

			' Add picture to the paragraph.
			par.Inlines.Add(pict1)

			' Save our document into DOCX format.
			dc.Save(documentPath)

			' Save our document into PDF format.
			dc.Save("PdfDocument.pdf")

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo("PdfDocument.pdf") With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
