using System;
using System.IO;
using System.Collections.Generic;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Sample
{
    class Sample
    {
        
        static void Main(string[] args)
        {
            ExtractPictures();
        }
		
        /// <summary>
        /// Extract all pictures from document (PDF, DOCX, RTF, HTML).
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/extract-pictures.php
        /// </remarks>
        public static void ExtractPictures()
        {
            // Path to a document where to extract pictures.
            string filePath = @"..\..\..\example.pdf";
           
            // Directory to store extracted pictures:
            DirectoryInfo imgDir = new DirectoryInfo("Extracted Pictures");
            imgDir.Create();
            string imgTemplateName = "Picture";

            // Here we store extracted images.
            List<ImageData> imgInventory = new List<ImageData>();

            // Load the document.
            DocumentCore dc = DocumentCore.Load(filePath);

            // Extract all images from document, skip duplicates.
            foreach (Picture pict in dc.GetChildElements(true, ElementType.Picture))
            {
                // Let's avoid the adding of duplicates.
                if (imgInventory.Exists((img => (img.GetStream().Length == pict.ImageData.GetStream().Length))) == false)
                    imgInventory.Add(pict.ImageData);
            }
            
            // Save all images.
            for (int i = 0; i < imgInventory.Count; i++)
            {
                string imagePath = Path.Combine(imgDir.FullName, String.Format("{0}{1}.{2}", imgTemplateName, i + 1, imgInventory[i].Format.ToString().ToLower()));
                File.WriteAllBytes(imagePath, imgInventory[i].GetStream().ToArray());                
            }

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(imgDir.FullName) { UseShellExecute = true });

        }
    }
}