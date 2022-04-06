using System;
using System.Text;
using SautinSoft.Document;
using SautinSoft.Document.CustomMarkups;
using SautinSoft.Document.Drawing;
using System.IO;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            InsertCheckBox();
        }
        /// <summary>
        /// Inserting a Check box content control.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/content-controls-insert-checkbox-net-csharp-vb.php
        /// </remarks>

        static void InsertCheckBox()
        {
            DocumentCore dc = new DocumentCore();

            InlineContentControl checkbox = new InlineContentControl(dc, ContentControlType.CheckBox);

            // Set the checkbox properties.
            checkbox.Properties.Title = "Click me";
            checkbox.Properties.Checked = true;
            checkbox.Properties.LockDeleting = true;
            checkbox.Properties.CharacterFormat.Size = 24;

            // Override default checkbox appearance.
            checkbox.Properties.CheckedSymbol.FontName = "Courier New";
            checkbox.Properties.CheckedSymbol.Character = 'X';
            checkbox.Properties.UncheckedSymbol.FontName = "Courier New";
            checkbox.Properties.UncheckedSymbol.Character = 'O';

            dc.Sections.Add(new Section(dc, new Paragraph(dc, new Run(dc, "Click me => ", 
                new CharacterFormat() {Size=24, FontColor=new Color("#3399FF") }), checkbox)));

            // Save our document into DOCX format.
            string resultPath = @"result.docx";
            dc.Save(resultPath, new DocxSaveOptions());

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}