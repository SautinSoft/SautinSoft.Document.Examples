using System;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System.IO;
using System.Linq;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            SecureDocument();
        }
		
        /// <summary>
        /// Create and secure a PDF document by password. Also set the permissions for the document.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/security-options-net-csharp-vb.php
        /// </remarks>
        public static void SecureDocument()
        {
            string filePath = @"ProtectedDocument.pdf";

            DocumentCore dc = new DocumentCore();

            // Let's create a simple document.
            dc.Content.End.Insert("Hello World!!!", new CharacterFormat() { FontName = "Verdana", Size = 65.5f, FontColor = Color.Orange });

            PdfSaveOptions so = new PdfSaveOptions();
            // Password Protection
            so.EncryptionDetails.UserPassword = "12345";
            // EncryptionAlgorithm
            so.EncryptionDetails.EncryptionAlgorithm = PdfEncryptionAlgorithm.RC4_128;
            //Permissions: Content Copying, Commenting, Printing, Changing the Document, filing of form fildes etc
            //Printing: Allowed
            so.EncryptionDetails.Permissions = PdfPermissions.Printing;
            
            // Save a document as the PDF file with Security Options.
            dc.Save(filePath, so);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}
