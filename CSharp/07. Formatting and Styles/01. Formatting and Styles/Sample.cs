using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            FormattingAndStyles();
        }
        /// <summary>
        /// Creates a new document and applies formatting and styles.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/formatting-and-styles.php
        /// </remarks>
        static void FormattingAndStyles()
        {
            string docxPath = @"FormattingAndStyles.docx";

            // Let's create a new document.
            DocumentCore dc = new DocumentCore();            
            Run run1 = new Run(dc, "This is Run 1 with character format Green. ");
            Run run2 = new Run(dc, "This is Run 2 with style Red.");
            
            // Create a new character style.
            CharacterStyle redStyle = new CharacterStyle("Red");
            redStyle.CharacterFormat.FontColor = Color.Red;
            dc.Styles.Add(redStyle);
            
            // Apply the direct character formatting.            
            run1.CharacterFormat.FontColor = Color.DarkGreen;

            // Apply only the style.
            run2.CharacterFormat.Style = redStyle;
            
            dc.Content.End.Insert(run1.Content);
            dc.Content.End.Insert(run2.Content);

            // Save our document into DOCX format.
            dc.Save(docxPath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(docxPath) { UseShellExecute = true });
        }
    }
}