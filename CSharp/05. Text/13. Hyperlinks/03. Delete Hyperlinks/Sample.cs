using System.Text;
using System.Linq;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            DeleteHyperlinksObjects();
            DeleteHyperlinksURL();
        }

        /// <summary>
        /// How to delete all hyperlink objects.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/hyperlinks-delete-url-csharp-vb-net.php
        /// </remarks>		
        public static void DeleteHyperlinksObjects()
        {
            // Let us say, we've a DOCX document.
            // And we've to remove the hyperlink objects.

            string inpFile = @"..\..\..\Hyperlinks example.docx";
            string outFile = @"Result - Delete Hyperlinks completely.pdf";
           
            // Let's open our document.
            DocumentCore dc = DocumentCore.Load(inpFile);

            // Loop by all hyperlinks and replace the URL (address).
            foreach (Hyperlink hpl in dc.GetChildElements(true, ElementType.Hyperlink).Reverse())
                hpl.ParentCollection.Remove(hpl);

            // Save our document back, but in PDF format.
            dc.Save(outFile, new PdfSaveOptions() { Compliance = PdfCompliance.PDF_14 });

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
        /// <summary>
        /// How to delete all hyperlinks but preserve only their text.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/hyperlinks-delete-url-csharp-vb-net.php
        /// </remarks>		
        public static void DeleteHyperlinksURL()
        {
            // Let us say, we've a DOCX document.
            // And we've to remove all hyperlinks, preserve only their text.

            // Note, we can't make the property 'Hyperlink.Address' empty, this is not allowed.
            // Therefore we have to remove the all 'Hyperlinks' object and 
            // insert the text objects 'Inline' instead of them.

            string inpFile = @"..\..\..\Hyperlinks example.docx";
            string outFile = @"Result - delete links and preserve text.docx";

            // Let's open our document.
            DocumentCore dc = DocumentCore.Load(inpFile);

            // Loop by all hyperlinks in a reverse, to remove the "Hyperlink" objects 
            // and replace them by their text ("Inline" objects).
            foreach (Hyperlink hpl in dc.GetChildElements(true, ElementType.Hyperlink).Reverse())
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
                    // Set the Auto font color (Black for the most cases) and remove the underline.
                    // Hide these lines if you want to preserve the formatting the same as the hyperlink had.
                    if (textInlines[i] is Run)
                    {
                        (textInlines[i] as Run).CharacterFormat.FontColor = Color.Auto;
                        (textInlines[i] as Run).CharacterFormat.UnderlineStyle = UnderlineType.None;
                    }
                    parentCollection.Insert(index + i, textInlines[i].Clone(true));
                }
            }
            // Save the document back.
            dc.Save(outFile);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}