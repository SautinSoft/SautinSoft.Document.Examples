using SautinSoft.Document;
using SautinSoft.Document.Tables;
using System.Linq;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            ImportStyles();
        }

        /// <summary>
        /// This sample shows how to import styles from a one document to another. 
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/styles-import.php
        /// </remarks>
        public static void ImportStyles()
        {
            DocumentCore source = DocumentCore.Load(@"..\..\..\document with styles.docx");
            DocumentCore dest = new DocumentCore();
            dest.Sections.Add(new Section(dest));

            // 1. Let's import all paragraph styles.
            foreach (Style s in source.Styles.Where(s => s is ParagraphStyle))
                dest.Styles.Import(s);

            // 1.5. Let's import Character style with name from "Green".

            // 1.5.1 Find a desired style with name "Green".
            CharacterStyle charStyle = (CharacterStyle)source.Styles.FirstOrDefault(s => s is CharacterStyle && s.Name=="Green");
            // 1.5.2 Import the style into the "dest" document and get the reference to it.
            if (charStyle != null)
                charStyle = (CharacterStyle)dest.Styles.Import(charStyle);

            // 2. Insert a new paragraph and apply the just now imported style.
            Paragraph p = new Paragraph(dest, "Charles Dickens was an extraordinary man. He is best known " +
                "as a novelist but he was very much more than that. He was as prominent in his other pursuits "+
                "but they were not areas of life where we can still see him today.");
            dest.Sections[0].Blocks.Add(p);

            // Find a style with the name "Center"
            ParagraphStyle pStyle = (ParagraphStyle)dest.Styles.FirstOrDefault(s => s is ParagraphStyle
                            && s.Name == "Center");
            
            // Apply the style to the paragraph.
            if (pStyle != null)
                p.ParagraphFormat.Style = pStyle;

            // Aplly the character style "Green" to all Run element in our paragraph.
            if (charStyle!=null)
            {
                foreach (Run r in p.Inlines)
                    r.CharacterFormat.Style = charStyle;
            }            

            // Save the dest document into DOCX format.
            string docPath = "SimpleImport.docx";
            dest.Save(docPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docPath) { UseShellExecute = true });
        }
    }
}