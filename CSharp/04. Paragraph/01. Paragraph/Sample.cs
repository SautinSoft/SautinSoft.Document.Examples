using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            Paragraph();
        }
        /// <summary>
        /// Creates a new document with a paragraph.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/paragraph.php
        /// </remarks>
        static void Paragraph()
        {
            DocumentCore dc = new DocumentCore();

            string filePath = @"Result-file.docx";

            dc.Sections.Add(
                new Section(dc,
                    new Paragraph(dc, "Text is right aligned.")
                    {
                        ParagraphFormat = new ParagraphFormat
                        {
                            Alignment = HorizontalAlignment.Right
                        }
                    },
                    new Paragraph(dc, "This paragraph has the following properties: Left indentation is 15 points, right indentation is 5 centimeters, hanging indentation is 10 points, line spacing is exactly 14 points, space before and space after are 10 points.")
                    {
                        ParagraphFormat = new ParagraphFormat
                        {
                            LeftIndentation = 15,
                            RightIndentation = LengthUnitConverter.Convert(5, LengthUnit.Centimeter, LengthUnit.Point),
                            SpecialIndentation = 10,
                            LineSpacing = 14,
                            LineSpacingRule = LineSpacingRule.Exactly,
                            SpaceBefore = 10,
                            SpaceAfter = 10
                        }
                    }));

            dc.Save(filePath);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}