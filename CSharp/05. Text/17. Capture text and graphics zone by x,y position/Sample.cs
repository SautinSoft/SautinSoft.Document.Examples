using System;
using System.IO;
using System.Linq;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using SautinSoft.Document.Tables;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free 100-day key here:   
            // https://sautinsoft.com/start-for-free/

            CaptureTextZoneByXY();
        }
        /// <summary>
        /// How to capture text and images from the existing PDF, DOCX, any document by specific (x,y) coordinates
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/capture-text-images-in-pdf-docx-document-by-specific-zone-x-y-coordinates-net-csharp-vb.php
        /// </remarks>
        static void CaptureTextZoneByXY()
        {
            // Let us say, we want to capture text and graphics into:            
            // Left-Top:(0, 50) mm,
            var mmXY = (0f, 50f);
            // Width: 250 mm, Height: 150 mm.
            var mmWH = (250f, 150f);

            // Zero-page index, e.g. page 1 has index 0.
            int[] pageCollection = new int[1] { 0 };

            // Convert mm to points
            double leftX = LengthUnitConverter.Convert(mmXY.Item1, LengthUnit.Millimeter, LengthUnit.Point);
            double topY = LengthUnitConverter.Convert(mmXY.Item2, LengthUnit.Millimeter, LengthUnit.Point);
            double width = LengthUnitConverter.Convert(mmWH.Item1, LengthUnit.Millimeter, LengthUnit.Point);
            double height = LengthUnitConverter.Convert(mmWH.Item2, LengthUnit.Millimeter, LengthUnit.Point);

            string inpFile = Path.GetFullPath(@"..\..\..\Potato Beetle.pdf");
            string outFile = "Result.docx";

            // 1. Load an existing document, load only specigic pages.
            PdfLoadOptions opt = new PdfLoadOptions()
            {
                SelectedPages = pageCollection,
                DetectTables = true,
                ConversionMode = PdfConversionMode.Flowing,
            };
            DocumentCore dc = DocumentCore.Load(inpFile, opt);

            // 2. Create new document to store captured data.
            DocumentCore dcCaptured = new DocumentCore();
            // Create import session.
            ImportSession session = new ImportSession(dc, dcCaptured, StyleImportingMode.KeepSourceFormatting);

            // 3. Iterate through document pages
            // and capture elements by (X,Y) and (width, height).
            var paginator = dc.GetPaginator(new PaginatorOptions() { UpdateFields = true });
            int pageIndex = 0;
            foreach (var page in paginator.Pages)
            {
                Section importedSection = null;
                foreach (var elementFrame in page.GetElementFrames(
                    ElementType.Paragraph,
                    ElementType.Table
                    ))
                {
                    // Is element inside capturing zone?
                    if (elementFrame.Bounds.Left >= leftX &&
                        elementFrame.Bounds.Left <= leftX + width &&
                        elementFrame.Bounds.Top >= topY &&
                        elementFrame.Bounds.Top <= topY + height)

                    {
                        if (importedSection == null)
                        {
                            importedSection = dcCaptured.Import<Section>(dc.Sections[pageIndex], false, session);
                            dcCaptured.Sections.Add(importedSection);
                        }

                        if (elementFrame.Element is Paragraph par)
                        {
                            var importedPar = dcCaptured.Import<Paragraph>(par, true, session);
                            importedSection.Blocks.Add(importedPar);
                        }
                        else if (elementFrame.Element is Table table)
                        {
                            var importedTable = dcCaptured.Import<Table>(table, true, session);
                            importedSection.Blocks.Add(importedTable);
                        }
                    }
                }
                pageIndex++;
            }
            dcCaptured.Save(outFile);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }        
    }
}