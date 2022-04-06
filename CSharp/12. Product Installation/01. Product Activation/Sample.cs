using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            ProductActivation();
        }
		
		/// <summary>
        /// Document .Net activation.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/product-activation.php
		/// </remarks>
        static void ProductActivation()
        {
            // Document .Net activation.
            
            // You will get own serial number after purchasing the license.
            // If you will have any questions, email us to sales@sautinsoft.com or ask at online chat https://www.sautinsoft.com.

            string serial = "1234567890";           

            // NOTICE: Place this line firstly, before creating of the DocumentCore object.
            DocumentCore.Serial = serial;

            // Let's create a new document by activated version.
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert("Hello World!", new CharacterFormat() { FontName = "Verdana", Size = 65.5f, FontColor = Color.Orange });

            // Save a document to a file in DOCX format.
            string filePath = @"Result.docx";
            dc.Save(filePath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}