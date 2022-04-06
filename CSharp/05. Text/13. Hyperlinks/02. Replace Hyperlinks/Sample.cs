using System.Text;
using System.Linq;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            ReplaceHyperlinksURL();
            ReplaceHyperlinksByText();
        }

        /// <summary>
        /// How to replace a hyperlink URL by a new address.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/hyperlinks-replace-url-csharp-vb-net.php
        /// </remarks>		
        public static void ReplaceHyperlinksURL()
        {
            // Let us say, we've a DOCX document.
            // And we've to replace the all URLs by the custom.
            // Furthermore, let's save the result as PDF.

            string inpFile = @"..\..\..\Hyperlinks example.docx";
            string outFile = @"Result - URL.pdf";

            // Let's open our document.
            DocumentCore dc = DocumentCore.Load(inpFile);

            // Specify the custom URL.
            string customURL = "https://www.sautinsoft.com";

            // Loop by all hyperlinks and replace the URL (address).
            foreach (Hyperlink hpl in dc.GetChildElements(true, ElementType.Hyperlink))
            {
                hpl.Address = customURL;
            }

            // Save our document back, but in PDF format.
            dc.Save(outFile, new PdfSaveOptions() { Compliance = PdfCompliance.PDF_14 });

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
        /// <summary>
        /// How to replace a hyperlink content and formatting. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/hyperlinks-replace-url-csharp-vb-net.php
        /// </remarks>		
        public static void ReplaceHyperlinksByText()
        {
            // Let us say, we've a DOCX document.
            // And we need to replace all hyperlinks by their text, color this text by red.
            // Also we have to preserve the rest formatting: font family, size and so on.

            string inpFile = @"..\..\..\Hyperlinks example.docx";
            string outFile = @"Result - Replace By Text.docx";

            // Let's open our document.
            DocumentCore dc = DocumentCore.Load(inpFile);

            // Loop by all hyperlinks in a reverse, to remove the "Hyperlink" objects 
            // and replace them by "Inline" objects.
            foreach (Hyperlink hpl in dc.GetChildElements(true, ElementType.Hyperlink).Reverse())
            {
                // Check that the Hyperlink is specified for a text element.
                if (hpl.DisplayInlines != null && hpl.DisplayInlines.Count > 0 && hpl.DisplayInlines[0] is Run)
                {
                    // Get the "Hyperlink" index in the parent collection.
                    InlineCollection parentCollection = hpl.ParentCollection;
                    int index = parentCollection.IndexOf(hpl);

                    // Get the "Hyperlink" text as the Inline collection.
                    InlineCollection textInlines = hpl.DisplayInlines;

                    // Remove the "Hyperlink" object from the parent collection by index.
                    parentCollection.RemoveAt(index);

                    // Insert the text (collection of Inlines) instead of the removed "Hyperlink" object 
                    // into the parent collection.
                    for (int i = 0; i < textInlines.Count; i++)
                    {
                        // Set the red font color, remove underline.
                        if (textInlines[i] is Run)
                        {
                            (textInlines[i] as Run).CharacterFormat.FontColor = Color.Red;
                            (textInlines[i] as Run).CharacterFormat.UnderlineStyle = UnderlineType.None;
                        }
                        parentCollection.Insert(index + i, textInlines[i].Clone(true));
                    }
                }
            }

            // Save our document back.
            dc.Save(outFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}