using SautinSoft.Document;
using System.IO;
using System.Text;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        { 
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            Unicode();
        }
        /// <summary>
        /// Convert files using different encodings.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/encodings.php
        /// </remarks>
        static void Unicode()
        {
            string inpFile = @"..\..\..\example.txt";
            string outFile = @"Result.pdf";

            // Provides access to an encoding provider for code pages that otherwise are available only in the desktop .NET Framework and .NET
            // https://learn.microsoft.com/en-us/dotnet/api/system.text.codepagesencodingprovider
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            DocumentCore dc = DocumentCore.Load(inpFile, new TxtLoadOptions { Encoding = Encoding.GetEncoding(1251)});
            dc.Save(outFile);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}