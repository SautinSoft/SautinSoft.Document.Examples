using SautinSoft.Document;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            CustomizedLists();
        }

        /// <summary>
        /// How to create customized lists.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-customized-list-in-docx-document-net-csharp-vb.php
        /// </remarks>		
        public static void CustomizedLists()
        {
            string documentPath = @"CustomizedLists.docx";

            // Let's create a new document.
            DocumentCore dc = new DocumentCore();

            // Add a new section.
            Section s = new Section(dc);
            dc.Sections.Add(s);

            // Create a new list style based on the ListTemplateType.Bullet template.
            ListStyle myCustomList1 = new ListStyle("MyCustomList1", ListTemplateType.Bullet);

            // Replace the standard bullet with the symbol ASTERISM (unicode 0x2042).
            myCustomList1.ListLevelFormats[0].NumberFormat = "\x2042";

            // Change the character formatting.
            myCustomList1.ListLevelFormats[0].CharacterFormat.FontName = "Calibri";
            myCustomList1.ListLevelFormats[0].CharacterFormat.FontColor = Color.Blue;
            dc.Styles.Add(myCustomList1);

            Paragraph a1 = new Paragraph(dc, "Asterism 1.");
            a1.ListFormat.Style = myCustomList1;
            s.Blocks.Add(a1);

            Paragraph a2 = new Paragraph(dc, "Asterism 2.");
            a2.ListFormat.Style = myCustomList1;
            a2.ParagraphFormat.SpaceAfter = 30;
            s.Blocks.Add(a2);

            // Create a new list style based on the ListTemplateType.NumberWithDot template.
            ListStyle myCustomList2 = new ListStyle("MyCustomList2", ListTemplateType.NumberWithDot);

            // Change the visual representation of the marker number on the first and second numbering levels.
            myCustomList2.ListLevelFormats[0].NumberFormat = "%1.";
            myCustomList2.ListLevelFormats[0].NumberStyle = NumberStyle.UpperRoman;
            myCustomList2.ListLevelFormats[0].TrailingCharacter = ListTrailingCharacter.Tab;
            myCustomList2.ListLevelFormats[1].NumberFormat = "%1. %2.";
            myCustomList2.ListLevelFormats[1].NumberStyle = NumberStyle.LowerLetter;
            myCustomList2.ListLevelFormats[1].TrailingCharacter = ListTrailingCharacter.Space;

            dc.Styles.Add(myCustomList2);

            Paragraph p1 = new Paragraph(dc, "First paragraph.");
            p1.ListFormat.Style = myCustomList2;
            s.Blocks.Add(p1);

            Paragraph p2 = new Paragraph(dc, "Second paragraph.");
            p2.ListFormat.Style = myCustomList2;
            s.Blocks.Add(p2);

            Paragraph p21 = new Paragraph(dc, "Sub paragraph a.");
            p21.ListFormat.Style = myCustomList2;
            p21.ListFormat.ListLevelNumber = 1;
            s.Blocks.Add(p21);

            Paragraph p22 = new Paragraph(dc, "Sub paragraph b.");
            p22.ListFormat.Style = myCustomList2;
            p22.ListFormat.ListLevelNumber = 1;
            s.Blocks.Add(p22);

            // Save our document into DOCX file.
            dc.Save(documentPath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(documentPath) { UseShellExecute = true });
        }
    }
}