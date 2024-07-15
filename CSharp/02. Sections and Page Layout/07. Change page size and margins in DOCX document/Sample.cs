using SautinSoft.Document;
using System.IO;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            // Get your free 100-day key here:   
            // https://sautinsoft.com/start-for-free/

            ChangePageProperties();
        }

        /// <summary>
        /// How to adjust a document page properties. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/page-setup.php
        /// </remarks>
        public static void ChangePageProperties()
        {
            var inpFile = Path.GetFullPath(@"..\..\..\example.docx");
            var outFile = Path.GetFullPath("Result.docx");
            DocumentCore dc = DocumentCore.Load(inpFile);

            // Apply page size, orientation and margins.
            foreach (var sect in dc.Sections)
            {
                sect.PageSetup.PaperType = PaperType.A3;
                sect.PageSetup.Orientation = Orientation.Landscape;
                sect.PageSetup.PageMargins.Left = LengthUnitConverter.Convert(5.0, LengthUnit.Centimeter, LengthUnit.Point);
            }

            dc.Save(outFile);

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
    }
}