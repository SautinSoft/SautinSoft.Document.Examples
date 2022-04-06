using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            AddAndDeleteParagraphs();
        }
        /// <summary>
        /// ElementCollection: Adds 20 paragraphs into document and delete 10 of them.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/elementcollection-add-delete.php
        /// </remarks>
        static void AddAndDeleteParagraphs()
        {
            DocumentCore dc = new DocumentCore();
            Section section = new Section(dc);
            dc.Sections.Add(section);
            for (int i = 0; i < 20; i++)
            {
                Paragraph par = new Paragraph(dc,"Text "+  i.ToString());
                section.Blocks.Add(par);
            }
            dc.Save("ResultFull.docx");
            for (int i = 0; i < section.Blocks.Count; )
            {
                section.Blocks.RemoveAt(i);
                i++;
            }
            dc.Save("ResultShort.docx");
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("ResultFull.docx") { UseShellExecute = true });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("ResultShort.docx") { UseShellExecute = true });
        }
    }
}