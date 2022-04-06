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
            InsertCombobox();
        }
        /// <summary>
        /// Inserting a Combo Box content control.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/content-controls-insert-combobox-net-csharp-vb.php
        /// </remarks>

        static void InsertCombobox()
        {
            // Let's create a simple document.
            DocumentCore dc = new DocumentCore();

            // Create a Combo Box content control.
            InlineContentControl combobox = new InlineContentControl(dc, ContentControlType.ComboBox);
            dc.Sections.Add(new Section(dc, new Paragraph(dc, new Run(dc, "Combo Box "), combobox)));

            // Set common the content control properties.
            combobox.Properties.Title = "Combo Box";
            combobox.Properties.LockDeleting = true;
            combobox.Properties.CharacterFormat.FontColor = Color.Blue;

            // Add combox's list items.
            combobox.Properties.ListItems.Add(new ContentControlListItem("One", "1"));
            combobox.Properties.ListItems.Add(new ContentControlListItem("Two", "2"));
            combobox.Properties.ListItems.Add(new ContentControlListItem("Three", "3"));
            combobox.Properties.ListItems.Add(new ContentControlListItem("Four", "4"));

            // Set a default selected item.
            combobox.Properties.SelectedListItem = combobox.Properties.ListItems[2]; 

            // Save our document into DOCX format.
            string resultPath = @"result.docx";
            dc.Save(resultPath, new DocxSaveOptions());
			
            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
    }
}