using SautinSoft.Document;
using System.Linq;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            FormDropDown();
        }
        /// <summary>
        /// Creates a document containing FormDropDown element.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/advanced.php
        /// </remarks>
        static void FormDropDown()
        {
            string filePath = @"Advanced.pdf";

            // Let's create document.
            DocumentCore dc = new DocumentCore();
            dc.Content.End.Insert(new Paragraph(dc, "The paragraph with FormDropDown element: ").Content);
            Paragraph par = dc.GetChildElements(true, ElementType.Paragraph).FirstOrDefault() as Paragraph;

            FormDropDownData field = new Field(dc, FieldType.FormDropDown).FormData as FormDropDownData;
            field.Items.Add("First Item");
            field.Items.Add("Second Item");
            field.Items.Add("Third Item");
            field.SelectedItemIndex = 2;

            par.Inlines.Add(field.Field);
            
            // Save our document.
            dc.Save(filePath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}