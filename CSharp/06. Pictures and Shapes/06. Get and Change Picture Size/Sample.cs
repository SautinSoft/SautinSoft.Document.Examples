using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Sample
{
    class Sample
    {
        
        static void Main(string[] args)
        {
            GetAndChangePictureSize();
        }

        /// <summary>
        /// Get a Picture size, Change it and Save the document back.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/get-and-change-picture-size-in-docx-csharp-vb-net.php
        /// </remarks>
        public static void GetAndChangePictureSize()
        {
            // Path to a document where to extract pictures.
            string inpFile = @"..\..\..\example.docx";
            string outFile = "Result.docx";

          
            // Load the document.
            DocumentCore dc = DocumentCore.Load(inpFile);

            // Get the physical size of the first picture from the document.            
            Picture pict = dc.GetChildElements(true, ElementType.Picture).FirstOrDefault() as Picture;
            Size size = pict.Layout.Size;

            Console.WriteLine("The 1st picture has this size:\r\n");
            Console.WriteLine("W: {0}, H: {1} (In points)", size.Width, size.Height);
            Console.WriteLine("W: {0}, H: {1} (In pixels)", LengthUnitConverter.Convert(size.Width, LengthUnit.Point, LengthUnit.Pixel),
                                                            LengthUnitConverter.Convert(size.Height, LengthUnit.Point, LengthUnit.Pixel));
            Console.WriteLine("W: {0:F2}, H: {1:F2} (In mm)", LengthUnitConverter.Convert(size.Width, LengthUnit.Point, LengthUnit.Millimeter),
                                                            LengthUnitConverter.Convert(size.Height, LengthUnit.Point, LengthUnit.Millimeter));
            Console.WriteLine("W: {0:F2}, H: {1:F2} (In cm)", LengthUnitConverter.Convert(size.Width, LengthUnit.Point, LengthUnit.Centimeter),
                                                            LengthUnitConverter.Convert(size.Height, LengthUnit.Point, LengthUnit.Centimeter));
            Console.WriteLine("W: {0:F2}, H: {1:F2} (In inches)", LengthUnitConverter.Convert(size.Width, LengthUnit.Point, LengthUnit.Inch),
                                                            LengthUnitConverter.Convert(size.Height, LengthUnit.Point, LengthUnit.Inch));

            Console.WriteLine("\r\nNow let\'s increase the picture size in x1.5 times. Press any key ...");
            Console.ReadKey();

            // Note, we don't change the physical picture size  we only scale/stretch the it.
            pict.Layout.Size = new Size(size.Width * 1.5, size.Height * 1.5);

            // Save the document as a new docx file.
            dc.Save(outFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });

        }
    }
}