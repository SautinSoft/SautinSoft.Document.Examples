using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            ConvertFromFile();
            ConvertFromStream();
        }

        /// <summary>
        /// Convert DOCX to DOCX (file to file).
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-docx-to-docx-in-csharp-vb.php
        /// </remarks>
        static void ConvertFromFile()
        {
            string inpFile = @"..\..\..\example.docx";
            string outFile = @"Result.docx";

            DocumentCore dc = DocumentCore.Load(inpFile);
            dc.Save(outFile);
			
			// Important for Linux: Install MS Fonts
			// sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }

        /// <summary>
        /// Convert DOCX to DOCX (using Stream).
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/convert-docx-to-docx-in-csharp-vb.php
        /// </remarks>
        static void ConvertFromStream()
        {

            // We need files only for demonstration purposes.
            // The conversion process will be done completely in memory.
            string inpFile = @"..\..\..\example.docx";
            string outFile = @"ResultStream.docx";
            byte[] inpData = File.ReadAllBytes(inpFile);
            byte[] outData = null;

            using (MemoryStream msInp = new MemoryStream(inpData))
            {

                // Load a document.
                DocumentCore dc = DocumentCore.Load(msInp, new DocxLoadOptions());

                // Save the document to DOCX format.
                using (MemoryStream outMs = new MemoryStream())
                {
                    dc.Save(outMs, new DocxSaveOptions() );
                    outData = outMs.ToArray(); 

					// Important for Linux: Install MS Fonts
					// sudo apt install ttf-mscorefonts-installer -y
                }
                // Show the result for demonstration purposes.
                if (outData != null)
                {
                    File.WriteAllBytes(outFile, outData);
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
                }
            }
        }
    }
}