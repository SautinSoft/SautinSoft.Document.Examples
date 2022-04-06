using System.IO;
using SautinSoft.Document;
using System;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            RecognizeImage();
        }

        /// <summary>
        /// Recognize an image using Tesseract (free OCR library) and save the result as DOCX document.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/ocr-image-using-tesseract-and-save-as-docx-net-csharp-vb.php
        /// </remarks>
        static void RecognizeImage()
        {
            // Here we'll recognize an image (perform OCR) containing a text on English, Russian and Vietnamese.
            // Next save the OCR result as a new DOCX document.

            // First steps:

            // 1. Download data files for English, Russian and Vietnamese languages.
            // Please download the files: eng.traineddata, rus.traineddata and vie.traineddata.
            // From here (good and fast): https://github.com/tesseract-ocr/tessdata_fast
            // or (best and slow): https://github.com/tesseract-ocr/tessdata_best

            // 2. Copy the files: eng.traineddata, rus.traineddata and vie.traineddata to
            // the folder "tessdata" in the Project root.

            // 3. Be sure that the folder "tessdata" also contains "pdf.ttf" file.

            // Let's start:
            string inpFile = @"..\..\..\image.png";
            string outFile = "Result.docx";

            ImageLoadOptions lo = new ImageLoadOptions();
            lo.OCROptions.OCRMode = OCRMode.Enabled;

            // You can specify all Tesseract parameters inside the method PerformOCR.
            lo.OCROptions.Method = PerformOCRTesseract;
            DocumentCore dc = DocumentCore.Load(inpFile, lo);

            // Make all text visible after Tesseract OCR (change font color to Black).
            // The matter is that Tesseract returns OCR result PDF document with invisible text.
            // But with help of Document .Net, we can change the text color, 
            // char scaling and spacing to desired.
            foreach (Run r in dc.GetChildElements(true, ElementType.Run))
            {
                r.CharacterFormat.FontColor = SautinSoft.Document.Color.Black;
                r.CharacterFormat.Scaling = 100;
                r.CharacterFormat.Spacing = 0;
                r.CharacterFormat.Size = 12;
            }

            // Change the page size and add page margins.
            Section section = dc.Sections[0];
            section.PageSetup.PaperType = PaperType.Letter;
            section.PageSetup.Orientation = Orientation.Landscape;
            double m = LengthUnitConverter.Convert(5, LengthUnit.Millimeter, LengthUnit.Point);
            section.PageSetup.PageMargins = new PageMargins() { Top = m, Left = m, Right = m, Bottom = m };
            dc.Save(outFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
        public static byte[] PerformOCRTesseract(byte[] image)
        {
            // Specify that Tesseract use three 3 languages: English, Russian and Vietnamese.
            string tesseractLanguages = "rus+eng+vie";


            // A path to a folder which contains languages data files and font file "pdf.ttf".
            // Language data files can be found here:
            // Good and fast: https://github.com/tesseract-ocr/tessdata_fast
            // or
            // Best and slow: https://github.com/tesseract-ocr/tessdata_best
            // Also this folder must have write permissions.
            string tesseractData = Path.GetFullPath(@"..\..\..\tessdata\");

            // A path for a temporary PDF file (because Tesseract returns OCR result as PDF document)
            string tempFile = Path.Combine(tesseractData, Path.GetRandomFileName());

            try
            {
                using (Tesseract.IResultRenderer renderer = Tesseract.PdfResultRenderer.CreatePdfRenderer(tempFile, tesseractData, true))
                {
                    using (renderer.BeginDocument("Serachablepdf"))
                    {
                        using (Tesseract.TesseractEngine engine = new Tesseract.TesseractEngine(tesseractData, tesseractLanguages, Tesseract.EngineMode.Default))
                        {
                            engine.DefaultPageSegMode = Tesseract.PageSegMode.Auto;
                            using (MemoryStream msImg = new MemoryStream(image))
                            {
                                System.Drawing.Image imgWithText = System.Drawing.Image.FromStream(msImg);
                                for (int i = 0; i < imgWithText.GetFrameCount(System.Drawing.Imaging.FrameDimension.Page); i++)
                                {
                                    imgWithText.SelectActiveFrame(System.Drawing.Imaging.FrameDimension.Page, i);
                                    using (MemoryStream ms = new MemoryStream())
                                    {
                                        imgWithText.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                        byte[] imgBytes = ms.ToArray();
                                        using (Tesseract.Pix img = Tesseract.Pix.LoadFromMemory(imgBytes))
                                        {
                                            using (var page = engine.Process(img, "Serachablepdf"))
                                            {
                                                renderer.AddPage(page);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                PdfLoadOptions pl = new PdfLoadOptions();
                pl.ShowInvisibleText = true;
				// 'Disabled' - Never load embedded fonts in PDF. Use the fonts with the same name installed at the system or similar by font metrics.
				// 'Enabled' - Always load embedded fonts in PDF.
				// 'Auto' - Load only embedded fonts missing in the system. In other case, use the system fonts.
                pl.PreserveEmbeddedFonts = PropertyState.Disabled;
                pl.ConversionMode = PdfConversionMode.Continuous;

                DocumentCore dc = DocumentCore.Load(File.OpenRead(tempFile + @".pdf"), pl);

                byte[] returnPdf;
                using (MemoryStream ms = new MemoryStream())
                {
                    PdfSaveOptions ps = new PdfSaveOptions();
                    dc.Save(ms, ps);
                    returnPdf = ms.ToArray();
                }
                return returnPdf;
            }
            catch (Exception e)
            {
                Console.WriteLine();
                Console.WriteLine("Please be sure that you have Language data files (*.traineddata) in your folder \"tessdata\"");
                Console.WriteLine("The Language data files can be download from here: https://github.com/tesseract-ocr/tessdata_fast");
                Console.ReadKey();
                throw new Exception("Error Tesseract: " + e.Message);
            }
            finally
            {
                if (File.Exists(tempFile + ".pdf"))
                    File.Delete(tempFile + ".pdf");
            }
        }
    }
}