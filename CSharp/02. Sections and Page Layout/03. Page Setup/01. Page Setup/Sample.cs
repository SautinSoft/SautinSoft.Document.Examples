using SautinSoft.Document;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            PageProperties();
        }
        
		/// <summary>
        /// How to adjust a document page properties. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/page-setup.php
        /// </remarks>
        public static void PageProperties()
        {
            string documentPath = @"PageProperties.docx";

            // Let's create a simple document.
            DocumentCore dc = new DocumentCore();

            // Add new section, B5 Landscape, and custom page margins.
            Section section1 = new Section(dc);
            section1.PageSetup.PaperType = PaperType.B5;
            section1.PageSetup.Orientation = Orientation.Landscape;
            section1.PageSetup.PageMargins = new PageMargins()
            {
                Top = LengthUnitConverter.Convert(50, LengthUnit.Millimeter, LengthUnit.Point),
                Right = LengthUnitConverter.Convert(1, LengthUnit.Inch, LengthUnit.Point),
                Bottom = LengthUnitConverter.Convert(10, LengthUnit.Millimeter, LengthUnit.Point),
                Left = LengthUnitConverter.Convert(2, LengthUnit.Centimeter, LengthUnit.Point)
            };

            dc.Sections.Add(section1);

            // Add some text to section1.
            section1.Content.Start.Insert("Shrek, a green ogre who loves the solitude in his swamp, " +
                            "finds his life interrupted when many fairytale characters are " +
                            "exiled there by order of the fairytale-hating Lord Farquaad.", new CharacterFormat() { FontName = "Times New Roman", Size = 14.0 });

            // Add page break.
            section1.Content.End.Insert(new SpecialCharacter(dc, SpecialCharacterType.PageBreak).Content);

            // Add another section, A4 Portrait, and custom page margins.
            Section section2 = new Section(dc);
            section2.PageSetup.PaperType = PaperType.A4;
            section2.PageSetup.Orientation = Orientation.Portrait;
            section2.PageSetup.PageMargins = new PageMargins()
            {
                Top = LengthUnitConverter.Convert(5, LengthUnit.Millimeter, LengthUnit.Point),
                Right = LengthUnitConverter.Convert(5, LengthUnit.Millimeter, LengthUnit.Point),
                Bottom = LengthUnitConverter.Convert(5, LengthUnit.Millimeter, LengthUnit.Point),
                Left = LengthUnitConverter.Convert(5, LengthUnit.Millimeter, LengthUnit.Point)
            };

            dc.Sections.Add(section2);

            // Add some text into section2.
            Paragraph p = new Paragraph(dc);
            p.Content.Start.Insert("Shrek tells them that he will go ask Farquaad to send them back. " +
                                    "He brings along a talking Donkey who is the only fairytale creature who knows the way to Duloc.",
                                    new CharacterFormat() { FontName = "Times New Roman", Size = 14.0 });
            p.ParagraphFormat.Alignment = HorizontalAlignment.Justify;
            section2.Blocks.Add(p);

            // Save our document into DOCX format.
            dc.Save(documentPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}