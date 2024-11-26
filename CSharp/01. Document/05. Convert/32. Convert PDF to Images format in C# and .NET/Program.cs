using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            ConvertPDFtoImages();
        }

        /// <summary>
        /// Convert PDF to Images (file to file).
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-pdf-to-images-in-csharp-vb.php
        /// </remarks>
        static void ConvertPDFtoImages()
        {
            // Path to a document where to extract pictures.
            // By the way: You may specify DOCX, HTML, RTF files also.
            DocumentCore dc = DocumentCore.Load(@"..\..\..\example.pdf");

            // PaginationOptions allow to know, how many pages we have in the document.
            DocumentPaginator dp = dc.GetPaginator(new PaginatorOptions());

            // Each document page will be saved in its own image format: PNG, JPEG, TIFF with different DPI.
            dp.Pages[0].Save(@"..\..\..\example.png", new ImageSaveOptions() { DpiX = 800, DpiY = 800 });
            dp.Pages[1].Save(@"..\..\..\example.jpeg", new ImageSaveOptions() { DpiX = 400, DpiY = 400 });
            dp.Pages[2].Save(@"..\..\..\example.tiff", new ImageSaveOptions() { DpiX = 650, DpiY = 650 });

        }
    }
}