using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            RasterizeDocxToPicture();
        }

        /// <summary>
        /// Rasterize DOCX document into PNG and JPEG.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/rasterize-docx-document-to-png-jpeg-net-csharp-vb.php
        /// </remarks>
        static void RasterizeDocxToPicture()
        {
            // In this example we'll how rasterize/convert 1st and 2nd pages of DOCX document
            // into PNG and JPEG images.
            string inputFile = @"..\..\..\example.docx";
            string jpegFile = @"..\..\..\Result.jpg";
            string pngFile = @"..\..\..\Result.png";
            // The file format is detected automatically from the file extension: ".docx".
            // But as shown in the example below, we can specify DocxLoadOptions as 2nd parameter
            // to explicitly set that a loadable document has Docx format.
            DocumentCore dc = DocumentCore.Load(inputFile, new DocxLoadOptions());

            DocumentPaginator documentPaginator = dc.GetPaginator(new PaginatorOptions() { UpdateFields = true });

            int pagesToRasterize = 2;
            int currentPage = 1;

            foreach (DocumentPage page in documentPaginator.Pages)
            {
                // Save the page to a file.
                if (currentPage == 1)
                    page.Save(pngFile, new ImageSaveOptions() { Format = ImageSaveFormat.Png });
                else if (currentPage == 2)
                    page.Save(jpegFile, new ImageSaveOptions() { Format = ImageSaveFormat.Jpeg });

                currentPage++;

                if (currentPage > pagesToRasterize)
                    break;
            }

            // Open the results for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pngFile) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(jpegFile) { UseShellExecute = true });
        }
    }
}
