using SautinSoft.Document;
using SautinSoft.Document.Tables;
using System.Linq;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            ImportKeepSourceFormatting();
        }

        /// <summary>
        /// Import an Element with Styles from another document. Mode: KeepSourceFormatting.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/import-element-keep-source-formatting.php
        /// </remarks>
        private static void ImportKeepSourceFormatting()
        {
            // Mode: KeepSourceFormatting.

            // 'KeepSourceFormatting' means to copy all required styles to the destination document, 
            // generate unique style names if needed.

            // For example, a destination document contains a style "Green" (Calibri, FontSize = 20, Green, Underline).
            // And a source document also contains an equal style with the same name "Green" (Calibri, FontSize = 20, Green, Underline).
            // The style "Green" will be imported and renamed to "Green1".
            // All imported elements linked to style "Green" will be remapped to style "Green1".

            DocumentCore source = DocumentCore.Load(@"..\..\..\SourceStyles.docx");
            DocumentCore dest = new DocumentCore();

            // Let's create a style "Green" (Calibri, FontSize = 20, Green, Underline).
            CharacterStyle chStyle = new CharacterStyle("Green");
            chStyle.CharacterFormat.FontName = "Calibri";
            chStyle.CharacterFormat.FontColor = Color.Green;
            chStyle.CharacterFormat.Size = 20;
            chStyle.CharacterFormat.UnderlineStyle = UnderlineType.Single;
            dest.Styles.Add(chStyle);
            dest.Content.End.Insert(new Run(dest, "This text has the style Green.", new CharacterFormat() { Style = chStyle }).Content);

            // Create an ImportSession with mode 'KeepSourceFormatting'.
            ImportSession session = new ImportSession(source, dest, StyleImportingMode.KeepSourceFormatting);

            // Let's import a paragraph.
            // The imported paragraph contains a text with style "Green" (FontSize = 20, Green, Underline). 
            // The style "Green" will be imported and renamed to "Green1", because we already have "Green".
            // All links in imported paragraph will be remapped to the style "Green1".
            Paragraph importedPar = dest.Import<Paragraph>((Paragraph)source.Sections[0].Blocks[0], true, session);
            dest.Content.End.Insert(importedPar.Content);                  

            // Save the destination document into DOCX format.
            string docPath = "KeepSourceFormatting.docx";
            dest.Save(docPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docPath) { UseShellExecute = true });
        }    
    }
}