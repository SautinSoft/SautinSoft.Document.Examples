using System;
using SautinSoft.Document;
using System.Text;
using SautinSoft.Document.Drawing;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/
            ParagraphAlignment();
        }
        /// <summary>
        /// Aligning paragraphs using tab stops: left, center, right.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/paragraph-alignment-tabstop.php
        /// </remarks>

        static void ParagraphAlignment()
      
        {
            string resultPath = @"..\..\..\Result.docx";
            string pictPath = @"..\..\..\logo.png";
            int SpecialIntendt = 35;

            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            ParagraphStyle paragraphStyle = new ParagraphStyle("ParagraphStyle1");
            paragraphStyle.CharacterFormat.FontName = "Arial Narrow";
            paragraphStyle.CharacterFormat.Size = 10;
            paragraphStyle.ParagraphFormat.Alignment = SautinSoft.Document.HorizontalAlignment.Left;
            paragraphStyle.ParagraphFormat.SpecialIndentation = SpecialIntendt;

            ParagraphStyle paragraphStyleJust = new ParagraphStyle("ParagraphStyle2");
            paragraphStyleJust.CharacterFormat.FontName = "Arial Narrow";
            paragraphStyleJust.CharacterFormat.Size = 10;
            paragraphStyleJust.ParagraphFormat.Alignment = SautinSoft.Document.HorizontalAlignment.Justify;

            ParagraphStyle paragraphStyleRight = new ParagraphStyle("ParagraphStyle3");
            paragraphStyleRight.CharacterFormat.FontName = "Arial Narrow";
            paragraphStyleRight.CharacterFormat.Size = 10;
            paragraphStyleRight.ParagraphFormat.Alignment = SautinSoft.Document.HorizontalAlignment.Right;

            CharacterStyle characterStyle = new CharacterStyle("CharacterStyle1");
            characterStyle.CharacterFormat.FontName = "Arial Narrow";
            characterStyle.CharacterFormat.UnderlineStyle = UnderlineType.Single;
            characterStyle.CharacterFormat.Size = 10;

            dc.Styles.Add(paragraphStyle);
            dc.Styles.Add(paragraphStyleJust);
            dc.Styles.Add(paragraphStyleRight);
            dc.Styles.Add(characterStyle);

            Picture pict2 = db.InsertImage(pictPath, new HorizontalPosition(-7.5, LengthUnit.Centimeter, HorizontalPositionAnchor.RightMargin),
                 new VerticalPosition(1, LengthUnit.Centimeter, VerticalPositionAnchor.TopMargin), WrappingStyle.InFrontOfText, new Size(65, 15, LengthUnit.Millimeter));

            // Paragraph 1.
            Paragraph p = new Paragraph(dc, "COMMUNITY ASSOCIATION AUTHORITY TO REPRESENT");

            p.ParagraphFormat.Alignment = SautinSoft.Document.HorizontalAlignment.Center;
            dc.Sections[0].Blocks.Add(p);

            // Paragraph 1.
            Paragraph p2 = new Paragraph(dc, "AND RETAINER AGREEMENT");
            p2.ParagraphFormat.Alignment = SautinSoft.Document.HorizontalAlignment.Center;
            dc.Sections[0].Blocks.Add(p2);

            Paragraph p3 = new Paragraph(dc, "This Agreement is effective from January 1, 2025 to December 31, 2029.")
            {
                ParagraphFormat = new ParagraphFormat
                {
                    Style = paragraphStyle
                }
            };
            dc.Sections[0].Blocks.Add(p3);

            Paragraph p4 = new Paragraph(dc, "This Agreement can be renewed for additional one year periods, AT THE ELECTION OF THE ASSOCIATION, by paying the Annual Retainer Fee for the next calendar year period.  The billing rate and the terms of this Agreement  effective for any renewal term remain the same, unless a change is announced by the Firm in writing.  This Agreement may be terminated at any time by either party by giving written notice to the other party.")
            {
                ParagraphFormat = new ParagraphFormat
                {
                    Style = paragraphStyle
                }
            };
            dc.Sections[0].Blocks.Add(p4);

            Paragraph p5 = new Paragraph(dc, "«Company__Also_Known_As_»   (“Association”),")
            {
                ParagraphFormat = new ParagraphFormat
                {
                    Alignment = SautinSoft.Document.HorizontalAlignment.Center
                }
            };
            dc.Sections[0].Blocks.Add(p5);

            Paragraph p6 = new Paragraph(dc, "by and through its Board of Directors, retains the law firm of TruckCom & Sautin, P.A. (hereinafter referred to as “Firm” or “Becker”), to represent it as legal counsel in the matters described in this Agreement and its Exhibits. This retainer and representation is solely and exclusively for the benefit of the Association, as a corporate entity, and not for any other party or third parties.  There are no intended third party beneficiaries, including, but not limited to members of the Association; residents living in the community operated by the Association or guests of such residents; officers, directors, employees, or agents of the Association.")
            {
                ParagraphFormat = new ParagraphFormat
                {
                    Style = paragraphStyleJust
                }
            };
            dc.Sections[0].Blocks.Add(p6);

            Paragraph p7 = new Paragraph(dc, "Paying the Annual Retainer Fee sum of Three Hundred Dollars ($300.00)  entitles the Association to the services listed in Exhibit “A” at no extra charge.")
            {
                ParagraphFormat = new ParagraphFormat
                {

                    Style = paragraphStyleJust
                }
            };
            dc.Sections[0].Blocks.Add(p7);

            Paragraph p8 = new Paragraph(dc, "The Firm will provide general legal services upon the request of the Association or as authorized by this Agreement concerning the day-to-day operation of the Association, including certain litigation, arbitration and mediation matters, on the reduced hourly fees stated in Exhibit “B” to this Agreement, subject to the terms and conditions in Exhibit “A”.")
            {
                ParagraphFormat = new ParagraphFormat
                {
                    Style = paragraphStyleJust
                }
            };
            dc.Sections[0].Blocks.Add(p8);

            Paragraph p9 = new Paragraph(dc, "The undersigned officer or agent of the Association represents and certifies that he(she) is authorized by the Board of Directors to execute this Agreement and agrees to the terms contained in the attached Exhibits on behalf of the Association and its membership.")
            {
                ParagraphFormat = new ParagraphFormat
                {
                    Style = paragraphStyleJust
                }
            };
            dc.Sections[0].Blocks.Add(p9);

            Paragraph p10 = new Paragraph(dc, "\"ASSOCIATION\"")
            {
                ParagraphFormat = new ParagraphFormat
                {
                    Style = paragraphStyleRight
                }
            };
            dc.Sections[0].Blocks.Add(p10);

            Paragraph p11 = new Paragraph(dc, "By:_________________________")
            {
                ParagraphFormat = new ParagraphFormat
                {
                    Style = paragraphStyleRight
                }
            };
            dc.Sections[0].Blocks.Add(p11);

            Paragraph p12 = new Paragraph(dc, "Print Nane:_________________________")
            {
                ParagraphFormat = new ParagraphFormat
                {
                    Style = paragraphStyleRight
                }
            };
            dc.Sections[0].Blocks.Add(p12);

            Paragraph p13 = new Paragraph(dc, "Print Title:_________________________")
            {
                ParagraphFormat = new ParagraphFormat
                {
                    Style = paragraphStyleRight
                }
            };
            dc.Sections[0].Blocks.Add(p13);

            Paragraph p114 = new Paragraph(dc, "Date:_________________________")
            {
                ParagraphFormat = new ParagraphFormat
                {
                    Style = paragraphStyleRight
                }
            };
            dc.Sections[0].Blocks.Add(p114);
           
            Paragraph p115 = new Paragraph(dc);
            p115.ParagraphFormat.Tabs.Add(new TabStop(350, TabStopAlignment.Left));
            p115.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.Tab));
            p115.Inlines.Add(new Run(dc, "By______________"));
            dc.Sections[0].Blocks.Add(p115);

            Paragraph p116 = new Paragraph(dc);
            p116.ParagraphFormat.Tabs.Add(new TabStop(350, TabStopAlignment.Left));
            p116.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.Tab));
            p116.Inlines.Add(new Run(dc, "Print Name__________"));
            dc.Sections[0].Blocks.Add(p116);

            Paragraph p117 = new Paragraph(dc);
            p117.ParagraphFormat.Tabs.Add(new TabStop(350, TabStopAlignment.Left));
            p117.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.Tab));
            p117.Inlines.Add(new Run(dc, "Print Title____________"));
            dc.Sections[0].Blocks.Add(p117);

            Paragraph p118 = new Paragraph(dc);
            p118.ParagraphFormat.Tabs.Add(new TabStop(350, TabStopAlignment.Left));
            p118.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.Tab));
            p118.Inlines.Add(new Run(dc, "Date___________"));
            dc.Sections[0].Blocks.Add(p118);

            Paragraph p119 = new Paragraph(dc);
            p119.ParagraphFormat.Tabs.Add(new TabStop(350, TabStopAlignment.Center));
            p119.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.Tab));
            p119.Inlines.Add(new Run(dc, "By______________"));
            dc.Sections[0].Blocks.Add(p119);

            Paragraph p120 = new Paragraph(dc);
            p120.ParagraphFormat.Tabs.Add(new TabStop(350, TabStopAlignment.Center));
            p120.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.Tab));
            p120.Inlines.Add(new Run(dc, "Print Name__________"));
            dc.Sections[0].Blocks.Add(p120);

            Paragraph p121 = new Paragraph(dc);
            p121.ParagraphFormat.Tabs.Add(new TabStop(350, TabStopAlignment.Center));
            p121.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.Tab));
            p121.Inlines.Add(new Run(dc, "Print Title____________"));
            dc.Sections[0].Blocks.Add(p121);

            Paragraph p122 = new Paragraph(dc);
            p122.ParagraphFormat.Tabs.Add(new TabStop(350, TabStopAlignment.Center));
            p122.Inlines.Add(new SpecialCharacter(dc, SpecialCharacterType.Tab));
            p122.Inlines.Add(new Run(dc, "Date___________"));
            dc.Sections[0].Blocks.Add(p122);

            // Save our document into DOCX format.
            dc.Save(resultPath, new DocxSaveOptions());
            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }

    }
}