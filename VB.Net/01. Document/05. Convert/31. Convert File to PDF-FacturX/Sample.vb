Imports System.IO
Imports SautinSoft.Document

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			ConvertFromFile()
			ConvertFromStream()
		End Sub

		''' <summary>
		''' Convert RTF file to PDF/ Factur-X format (file to file).
		''' Read more information about Factur-X: https://fnfe-mpe.org/factur-x/
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/convert-file-to-pdf-factur-x-in-csharp-vb.php
		''' </remarks>
		Private Shared Sub ConvertFromFile()
			Dim inpFile As String = "..\..\..\Factur.rtf"
			Dim outFile As String = "..\..\..\Factur.pdf"
			Dim xmlInfo As String = File.ReadAllText("..\..\..\Factur\Facture.xml")

			Dim dc As DocumentCore = DocumentCore.Load(inpFile)

			Dim pdfSO As New PdfSaveOptions() With {.FacturXXML = xmlInfo}

			dc.Save(outFile, pdfSO)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub

		''' <summary>
		''' Convert PDF file to PDF/ Factur-X format (using Stream).
		''' Read more information about Factur-X: https://fnfe-mpe.org/factur-x/
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/convert-file-to-pdf-factur-x-in-csharp-vb.php
		''' </remarks>
		Private Shared Sub ConvertFromStream()

			' We need files only for demonstration purposes.
			' The conversion process will be done completely in memory.
			Dim inpFile As String = "..\..\..\Sample.pdf"
			Dim outFile As String = "..\..\..\Factur.pdf"
			Dim xmlInfo As String = File.ReadAllText("..\..\..\Factur\Facture.xml")
			Dim inpData() As Byte = File.ReadAllBytes(inpFile)
			Dim outData() As Byte = Nothing

			Using msInp As New MemoryStream(inpData)
				' Specifying PdfLoadOptions we explicitly set that a loadable document is PDF.
				Dim pdfLO As New PdfLoadOptions() With {
					.RasterizeVectorGraphics = False,
					.DetectTables = False,
					.PreserveEmbeddedFonts = PropertyState.Auto
				}

				' Load a document.
				Dim dc As DocumentCore = DocumentCore.Load(msInp, pdfLO)

				' Save the document to PDF/A format.
				Dim pdfSO As New PdfSaveOptions() With {.FacturXXML = xmlInfo}

				Using outMs As New MemoryStream()
					dc.Save(outMs, pdfSO)
					outData = outMs.ToArray()
				End Using
				' Show the result for demonstration purposes.
				If outData IsNot Nothing Then
					File.WriteAllBytes(outFile, outData)
					System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
				End If
			End Using
		End Sub
	End Class
End Namespace
