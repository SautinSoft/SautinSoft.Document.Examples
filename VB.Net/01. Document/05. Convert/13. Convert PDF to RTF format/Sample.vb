Imports System.IO
Imports SautinSoft.Document

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			ConvertFromFile()
			ConvertFromStream()
		End Sub
                ''' Get your free trial key here:   
                ''' https://sautinsoft.com/start-for-free/
		''' <summary>
		''' Convert PDF to RTF (file to file).
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-pdf-to-rtf-in-csharp-vb.php
		''' </remarks>
		Private Shared Sub ConvertFromFile()
            Dim inpFile As String = "..\..\..\example.pdf"
            Dim outFile As String = "Result.rtf"

            ' Specifying PdfLoadOptions we explicitly set that a loadable document is PDF.
            Dim pdfLO As New PdfLoadOptions()
            With pdfLO
                .RasterizeVectorGraphics = False
                .DetectTables = False
                ' 'Disabled' - Never load embedded fonts in PDF. Use the fonts with the same name installed at the system or similar by font metrics.
                ' 'Enabled' - Always load embedded fonts in PDF.
                ' 'Auto' - Load only embedded fonts missing in the system. In other case, use the system fonts.
                .PreserveEmbeddedFonts = PropertyState.Auto
            End With

            ' RasterizeVectorGraphics = False
            ' This means to load vector graphics as is. Don't transform it to raster images.

            ' DetectTables = False
            ' This means don't detect tables.
            ' The PDF format doesn't have real tables, in fact it's a set of orthogonal graphic lines.
            ' Set it to 'True' and the component will detect and recreate tables from graphic lines.

            Dim dc As DocumentCore = DocumentCore.Load(inpFile, pdfLO)
            dc.Save(outFile)
			
			' Important for Linux: Install MS Fonts
			' sudo apt install ttf-mscorefonts-installer -y

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
		End Sub

		''' <summary>
		''' Convert PDF to RTF (using Stream).
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-pdf-to-rtf-in-csharp-vb.php
		''' </remarks>
		Private Shared Sub ConvertFromStream()

            ' We need files only for demonstration purposes.
            ' The conversion process will be done completely in memory.
            Dim inpFile As String = "..\..\..\example.pdf"
            Dim outFile As String = "ResultStream.rtf"
			Dim inpData() As Byte = File.ReadAllBytes(inpFile)
			Dim outData() As Byte = Nothing

            Using msInp As New MemoryStream(inpData)

                ' Specifying PdfLoadOptions we explicitly set that a loadable document is PDF.
                Dim pdfLO As New PdfLoadOptions()
                With pdfLO
                    .RasterizeVectorGraphics = False
                    .DetectTables = False
                    ' 'Disabled' - Never load embedded fonts in PDF. Use the fonts with the same name installed at the system or similar by font metrics.
                    ' 'Enabled' - Always load embedded fonts in PDF.
                    ' 'Auto' - Load only embedded fonts missing in the system. In other case, use the system fonts.					
                    .PreserveEmbeddedFonts = PropertyState.Auto
                End With

                ' RasterizeVectorGraphics = False
                ' This means to load vector graphics as is. Don't transform it to raster images.

                ' DetectTables = False
                ' This means don't detect tables.
                ' The PDF format doesn't have real tables, in fact it's a set of orthogonal graphic lines.
                ' Set it to 'True' and the component will detect and recreate tables from graphic lines.

                ' Load a document.
                Dim dc As DocumentCore = DocumentCore.Load(msInp, pdfLO)

                ' Save the document to RTF format.
                Using outMs As New MemoryStream()
                    dc.Save(outMs, New RtfSaveOptions())
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
