using System;
using SautinSoft.Document;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            ParagraphFormatting();
        }

		/// <summary>
        /// This sample shows how to specify a paragraph formatting. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/paragraph-format.php
        public static void ParagraphFormatting()
        {
            string documentPath = @"ParagraphFormatting.docx";

            // Let's create a simple document.
            DocumentCore dc = new DocumentCore();

            // Add new section
            Section s = new Section(dc);
            dc.Sections.Add(s);

            // Paragraph 1.
            Paragraph p = new Paragraph(dc, "First paragraph!");
            dc.Sections[0].Blocks.Add(p);

            p.ParagraphFormat.Alignment = HorizontalAlignment.Left;
            p.ParagraphFormat.SpaceAfter = 20.0;
            p.ParagraphFormat.Borders.Add(MultipleBorderTypes.All, BorderStyle.Single, Color.Orange, 2f);

            // Paragraph 2.
            dc.Content.End.Insert("\nThis is a second paragraph.", new CharacterFormat() { FontName = "Calibri", Size = 16.0, FontColor = Color.Black, Bold = true });
            (dc.Sections[0].Blocks[1] as Paragraph).ParagraphFormat.Alignment = HorizontalAlignment.Center;

            // Create multiple paragraphs.
            for (int i = 0; i < 10; i++)
            {
                Paragraph pN = new Paragraph(dc, String.Format("Paragraph {0}. ", i + 1));
                dc.Sections[0].Blocks.Add(pN);

                pN.Content.End.Insert(new SpecialCharacter(dc, SpecialCharacterType.Tab).Content);
                Run run = new Run(dc, "Hello!");
                run.CharacterFormat.FontColor = Color.White;
                pN.Content.End.Insert(run.Content);

                pN.ParagraphFormat.BackgroundColor = new Color((int)(0xFF358CCB * (i + 1)));
                pN.ParagraphFormat.SpaceBefore = LengthUnitConverter.Convert(1, LengthUnit.Centimeter, LengthUnit.Point);
                pN.ParagraphFormat.SpaceAfter = LengthUnitConverter.Convert(10, LengthUnit.Millimeter, LengthUnit.Point);
            }

            // Save our document into DOCX format.
            dc.Save(documentPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}