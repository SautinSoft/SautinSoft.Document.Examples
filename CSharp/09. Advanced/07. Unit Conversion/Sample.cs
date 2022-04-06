using System;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            ConvertPointsTo();
            UsingUnitConversion();
        }

        /// <summary>
        /// Shows what is the one point is equal in different units.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/unit-conversion.php
        public static void ConvertPointsTo()
        {
            foreach (LengthUnit unit in Enum.GetValues(typeof(LengthUnit)))
            {
                string s = string.Format("1 point = {0} {1}",
                    LengthUnitConverter.Convert(1, LengthUnit.Point, unit),
                    unit.ToString().ToLowerInvariant());

                Console.WriteLine(s);                
            }
            Console.WriteLine("Press any key...");
            Console.ReadKey();
        }

        /// <summary>
        /// How to convert between different units.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/unit-conversion.php
        public static void UsingUnitConversion()
        {
            DocumentCore dc = new DocumentCore();

            // Add new section.
            Section s1 = new Section(dc);
            s1.PageSetup.PageWidth = LengthUnitConverter.Convert(100, LengthUnit.Millimeter, LengthUnit.Point);
            s1.PageSetup.PageHeight = LengthUnitConverter.Convert(1, LengthUnit.Inch, LengthUnit.Point);
            s1.PageSetup.PageMargins.Left = LengthUnitConverter.Convert(0.3, LengthUnit.Centimeter, LengthUnit.Point);
            s1.PageSetup.PageMargins.Top = LengthUnitConverter.Convert(7, LengthUnit.Pixel, LengthUnit.Point);
            s1.PageSetup.PageMargins.Right = LengthUnitConverter.Convert(100, LengthUnit.Twip, LengthUnit.Point);
            s1.PageSetup.PageMargins.Bottom = LengthUnitConverter.Convert(5, LengthUnit.Millimeter, LengthUnit.Point);

            dc.Sections.Add(s1);
            s1.Content.Start.Insert("You are welcome!", new CharacterFormat() { FontColor = Color.Orange, Size = 36 });

            dc.Save("Result.docx");

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("Result.docx") { UseShellExecute = true });
        }
    }
}