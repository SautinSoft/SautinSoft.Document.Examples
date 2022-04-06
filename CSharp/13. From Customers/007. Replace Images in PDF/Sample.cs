using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using SautinSoft.Document.Tables;
using System.IO;
namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            ReplaceImagesInPdf();
        }

        /// <summary>
        /// How to replace images in PDF document.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-replace-images-in-pdf-in-csharp-vb-net.php
        /// </remarks>
        static void ReplaceImagesInPdf()
        {

            // Path to a loadable document.
            string loadPath = @"..\..\..\example.pdf";
            string pictPath = @"..\..\..\replaceNA.jpg";

            // Load a document intoDocumentCore.
            DocumentCore dc = DocumentCore.Load(loadPath);
            
            // Load the Picture from a file.
            Picture picture = new Picture(dc, InlineLayout.Inline(new Size()), pictPath);
            
            // Find all pictures in the document.
            foreach (Element el in dc.GetChildElements(true, ElementType.Picture).Reverse())
            {
                if (el is Picture)
                {
                    // Ñopy all properties of the found picture and assign these properties to the new picture.
                    // If you do not do this, the picture may be inserted into an arbitrary place in the document. 
                    if (((Picture)el).Layout is FloatingLayout)
                    {
                        FloatingLayout old = (FloatingLayout )((Picture)el).Layout;
                        picture.Layout = FloatingLayout.Floating(old.HorizontalPosition, old.VerticalPosition, old.Size);
                    }

                    // Replace picture.
                    el.Content.Replace(picture.Content);
                }
            }

            // Save our document into PDF format.
            string savePath = @"replaced.pdf";
            dc.Save(savePath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(loadPath) { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(savePath) { UseShellExecute = true });
        }
    }
}