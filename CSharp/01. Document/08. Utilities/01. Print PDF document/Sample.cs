using SautinSoft.Document;
using System;
using System.IO;
using System.Text;
using System.Diagnostics;

namespace Example
{
    class Program
    {       
        static void Main(string[] args)
        {
            CreateAndPrintPdf();
        }

        /// <summary>
        /// Creates a new document and saves it into PDF format. Print PDF using your default printer.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/print-document-net-csharp-vb.php
        /// </remarks>
        public static void CreateAndPrintPdf()
        {
            // Set the path to create pdf document.
            string pdfPath = @"Result.pdf";

            // Let's create a simple PDF document.
            DocumentCore dc = new DocumentCore();

            // Add new section.
            Section section = new Section(dc);
            dc.Sections.Add(section);

            // Let's set page size A4.
            section.PageSetup.PaperType = PaperType.A4;

            // Add a paragraph using ContentRange:
            dc.Content.End.Insert("\nHi Dear Friends.", new CharacterFormat() { Size = 25, FontColor = Color.Blue, Bold = true });
            SpecialCharacter lBr = new SpecialCharacter(dc, SpecialCharacterType.LineBreak);
            dc.Content.End.Insert(lBr.Content);
            dc.Content.End.Insert("I'm happy to see you!", new CharacterFormat() { Size = 20, FontColor = Color.DarkGreen, UnderlineStyle = UnderlineType.Single });

            // Save PDF to a file
            dc.Save(pdfPath, new PdfSaveOptions());

            // Create a new process: Acrobat Reader. You may change in on Foxit Reader.
            string processFilename = Microsoft.Win32.Registry.LocalMachine
                     .OpenSubKey("Software")
                     .OpenSubKey("Microsoft")
                     .OpenSubKey("Windows")
                     .OpenSubKey("CurrentVersion")
                     .OpenSubKey("App Paths")
                     .OpenSubKey("Acrobat.exe")
                     .GetValue(String.Empty).ToString();

            // Let's transfer our PDF file to the process Adobe Reader
            ProcessStartInfo info = new ProcessStartInfo();
            info.Verb = "print";
            info.FileName = processFilename;
            info.Arguments = String.Format("/p /h {0}", pdfPath);
            info.CreateNoWindow = true;
            
            //(It won't be hidden anyway... thanks Adobe!)
            info.WindowStyle = ProcessWindowStyle.Hidden;
            info.UseShellExecute = false;

            Process p = Process.Start(info);
            p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

            // Print our PDF
            int counter = 0;
            while (!p.HasExited)
            {
                System.Threading.Thread.Sleep(1000);
                counter += 1;
                if (counter == 5) break;
            }
            if (!p.HasExited)
            {
                p.CloseMainWindow();
                p.Kill();
            }

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(pdfPath) { UseShellExecute = true });
        }
    }
}