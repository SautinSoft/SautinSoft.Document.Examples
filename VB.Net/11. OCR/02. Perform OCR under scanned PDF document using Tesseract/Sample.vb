Imports System.IO
Imports SautinSoft.Document
Imports System

Namespace Example
    Friend Class Program
        Public Shared Sub Main(args As String())
            ' Get your free trial key here:   
            ' https://sautinsoft.com/start-for-free/

            Call LoadScannedPdf()
        End Sub

        ''' <summary>
        ''' Load a scanned PDF document with help of Tesseract OCR (free OCR library) and save the result as DOCX document.
        ''' </summary>
        ''' <remarks>
        ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/ocr-load-scanned-pdf-using-tesseract-and-save-as-docx-net-csharp-vb.php
        ''' </remarks>
        Private Shared Sub LoadScannedPdf()
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
            Dim inpFile = "..\..\..\scan.pdf"
            Dim outFile = "Result.docx"

            Dim lo As PdfLoadOptions = New PdfLoadOptions()
            lo.OCROptions.OCRMode = OCRMode.Enabled
            ' 'false' - Never load embedded fonts in PDF. Use the fonts with the same name installed at the system or similar by font metrics.
            ' 'true' - Always load embedded fonts in PDF.
            lo.PreserveEmbeddedFonts = True

            ' You can specify all Tesseract parameters inside the method PerformOCR.
            lo.OCROptions.Method = AddressOf PerformOCRTesseract
            Dim dc = DocumentCore.Load(inpFile, lo)

            ' Make all text visible after Tesseract OCR (change font color to Black).
            ' The matter is that Tesseract returns OCR result PDF document with invisible text.
            ' But with help of Document .Net, we can change the text color, 
            ' char scaling and spacing to desired.
            For Each r As Run In dc.GetChildElements(True, ElementType.Run)
                r.CharacterFormat.FontColor = Color.Black
                r.CharacterFormat.Scaling = 100
                r.CharacterFormat.Spacing = 0
            Next

            dc.Save(outFile)

            ' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(outFile) With {.UseShellExecute = True})
        End Sub
        Public Shared Function PerformOCRTesseract(image As Byte()) As Byte()
            ' Specify that Tesseract use three 3 languages: English, Russian and Vietnamese.
            Dim tesseractLanguages = "eng+rus+vie"

            ' A path to a folder which contains languages data files and font file "pdf.ttf".
            ' Language data files can be found here:
            ' Good and fast: https://github.com/tesseract-ocr/tessdata_fast
            ' or
            ' Best and slow: https://github.com/tesseract-ocr/tessdata_best
            ' Also this folder must have write permissions.
            Dim tesseractData = Path.GetFullPath("..\..\..\tessdata\")

            ' A path for a temporary PDF file (because Tesseract returns OCR result as PDF document)
            Dim tempFile As String = Path.Combine(tesseractData, Path.GetRandomFileName())

            Try
                Using renderer = Tesseract.ResultRenderer.CreatePdfRenderer(tempFile, tesseractData, True)
                    Using renderer.BeginDocument("Serachablepdf")
                        Using engine As Tesseract.TesseractEngine = New Tesseract.TesseractEngine(tesseractData, tesseractLanguages, Tesseract.EngineMode.Default)
                            engine.DefaultPageSegMode = Tesseract.PageSegMode.Auto

                            Dim imgBytes = image
                            Using img = Tesseract.Pix.LoadFromMemory(imgBytes)
                                Using page = engine.Process(img, "Serachablepdf")
                                    renderer.AddPage(page)
                                End Using
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

            End Try
        End Function
    End Class
End Namespace
