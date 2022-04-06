using SautinSoft.Document;
using SautinSoft.Document.Tables;
using System.Linq;

namespace Sample
{
    class Sample
    {

        static void Main(string[] args)
        {
            ImportUseDestinationStyles();
        }

        /// <summary>
        /// Import an Element with Styles from another document. Mode: UseDestinationStyles.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/import-element-use-destination-styles.php
        /// </remarks>
        private static void ImportUseDestinationStyles()
        {
            // Mode: UseDestinationStyles.

            // 'UseDestinationStyles' means to copy only styles wchich are don't exist 
            // in the destination document. 
            // If the destination document already contains a style with the same name, 
            // therefore a style from a source will not be copied.

            // For example, a destination document contains a style "Green" (FontSize = 24, DarkGreen).
            // And a source document also contains a style with name "Green" (FontSize = 20, Green, Underline).
            // After the importing, the imported content will change its formatting correspondly to the "Green" style in the destination document.
            // Because style "Green" (FontSize = 20, Green, Underline) was not imported.

            DocumentCore source = DocumentCore.Load(@"..\..\..\SourceStyles.docx");
            DocumentCore dest = new DocumentCore();

            // Before importing a style from another document, let's create a style 
            // with the same name but different formatting to see 
            // how the name conflict will be resolved in mode 'UseDestinationStyles'.
            CharacterStyle chStyle = new CharacterStyle("Green");
            chStyle.CharacterFormat.FontColor = Color.DarkGreen;
            chStyle.CharacterFormat.Size = 24;
            dest.Styles.Add(chStyle);
            dest.Content.End.Insert(new Run(dest, "First ", new CharacterFormat() { Style = chStyle }).Content);
            
            // Create an ImportSession with mode 'UseDestinationStyles'.
            ImportSession session = new ImportSession(source, dest, StyleImportingMode.UseDestinationStyles);

            // Let's import a 1st paragraph from the source document.
            // The paragraph contains a text marked by style "Green". 
            // As a style with the same name is already exist, the new "Green" style will not be imported.
            Paragraph importedPar = dest.Import<Paragraph>((Paragraph)source.Sections[0].Blocks[0], true, session);
            dest.Content.End.Insert(importedPar.Content);                  

            // Save the destination document into DOCX format.
            string docPath = "UseDestinationStyles.docx";
            dest.Save(docPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docPath) { UseShellExecute = true });
        }    
    }
}