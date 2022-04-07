Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        LoadScannedPdf()
    End Sub

    ''' <summary>
    ''' Load a scanned PDF document with help of Tesseract OCR (free OCR library) and save the result as DOCX document.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/ocr-load-scanned-pdf-using-tesseract-and-save-as-docx-net-csharp-vb.php
    ''' </remarks>
    Private Sub LoadScannedPdf()
		' Here we'll load a scanned PDF document (perform OCR) containing a text on English, Russian and Vietnamese.
		' Next save the OCR result as a new DOCX document.

		' First steps:

		' 1. Download data files for English, Russian and Vietnamese languages.
		' Please download the files: eng.traineddata, rus.traineddata and vie.traineddata.
		' From here (good and fast): https://github.com/tesseract-ocr/tessdata_fast
		' or (best and slow): https://github.com/tesseract-ocr/tessdata_best

		' 2. Copy the files: eng.traineddata, rus.traineddata and vie.traineddata to
		' the folder "tessdata" in the Project root.

		' 3. Be sure that the folder "tessdata" also contains "pdf.ttf" file.

		' Let's start:
		Dim inpFile As String = "..\..\..\scan.pdf"
		Dim outFile As String = "Result.docx"

        Dim lo As New PdfLoadOptions()
        lo.OCROptions.OCRMode = OCRMode.Enabled
		' 'Disabled' - Never load embedded fonts in PDF. Use the fonts with the same name installed at the system or similar by font metrics.
		' 'Enabled' - Always load embedded fonts in PDF.
		' 'Auto' - Load only embedded fonts missing in the system. In other case, use the system fonts.		
        lo.PreserveEmbeddedFonts = PropertyState.Enabled

		' You can specify all Tesseract parameters inside the method PerformOCR.
		lo.OCROptions.Method = AddressOf PerformOCRTesseract
		Dim dc As DocumentCore = DocumentCore.Load(inpFile, lo)

        'Make all text visible after Tesseract OCR (change font color to Black).
        'The matter is that Tesseract returns OCR result PDF document with invisible text.
        'but with help of Document .Net, we can change the text color, 
        'char scaling And spacing to desired.
        For Each r As Run In dc.GetChildElements(True, ElementType.Run)
            r.CharacterFormat.FontColor = SautinSoft.Document.Color.Black
			r.CharacterFormat.Scaling = 100
			r.CharacterFormat.Spacing = 0
		Next r

        dc.Save(outFile)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
    End Sub
	Public Function PerformOCRTesseract(ByVal image() As Byte) As Byte()
		' Specify that Tesseract use three 3 languages: English, Russian and Vietnamese.
		'string tesseractLanguages = "rus+eng+vie";
		Dim tesseractLanguages As String = "eng"

		' A path to a folder which contains languages data files and font file "pdf.ttf".
		' Language data files can be found here:
		' Good and fast: https://github.com/tesseract-ocr/tessdata_fast
		' or
		' Best and slow: https://github.com/tesseract-ocr/tessdata_best
		' Also this folder must have write permissions.
		Dim tesseractData As String = Path.GetFullPath("..\..\..\tessdata\")

		' A path for a temporary PDF file (because Tesseract returns OCR result as PDF document)
		Dim tempFile As String = Path.Combine(tesseractData, Path.GetRandomFileName())

		Try
			Using renderer As Tesseract.IResultRenderer = Tesseract.PdfResultRenderer.CreatePdfRenderer(tempFile, tesseractData, True)
				Using renderer.BeginDocument("Serachablepdf")
					Using engine As New Tesseract.TesseractEngine(tesseractData, tesseractLanguages, Tesseract.EngineMode.Default)
						engine.DefaultPageSegMode = Tesseract.PageSegMode.Auto
						Using msImg As New MemoryStream(image)
							Dim imgWithText As System.Drawing.Image = System.Drawing.Image.FromStream(msImg)
							Dim i As Integer = 0
							Do While i < imgWithText.GetFrameCount(System.Drawing.Imaging.FrameDimension.Page)
								imgWithText.SelectActiveFrame(System.Drawing.Imaging.FrameDimension.Page, i)
								Using ms As New MemoryStream()
									imgWithText.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
									Dim imgBytes() As Byte = ms.ToArray()
									Using img As Tesseract.Pix = Tesseract.Pix.LoadFromMemory(imgBytes)
#Disable Warning BC42020 ' Variable declaration without an 'As' clause
										Using page = engine.Process(img, "Serachablepdf")
#Enable Warning BC42020 ' Variable declaration without an 'As' clause
											renderer.AddPage(page)
										End Using
									End Using
								End Using
								i += 1
							Loop
						End Using
					End Using
				End Using
			End Using

			Return File.ReadAllBytes(tempFile & ".pdf")
		Catch e As Exception
			Console.WriteLine()
			Console.WriteLine("Please be sure that you have Language data files (*.traineddata) in your folder ""tessdata""")
			Console.WriteLine("The Language data files can be download from here: https://github.com/tesseract-ocr/tessdata_fast")
			Console.ReadKey()
			Throw New Exception("Error Tesseract: " & e.Message)
		Finally
			If File.Exists(tempFile & ".pdf") Then
				File.Delete(tempFile & ".pdf")
			End If
		End Try
	End Function
End Module