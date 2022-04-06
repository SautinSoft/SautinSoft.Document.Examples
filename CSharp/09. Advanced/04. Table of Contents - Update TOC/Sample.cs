using System;
using System.IO;
using System.Linq;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;

namespace Example
{
    class Sample
    {
        static void Main(string[] args)
        {
            TOC();
        }
		
		/// <summary>
        /// Update table of contents in word document.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/update-table-of-contents-in-word-document-net-csharp-vb.php
        /// </remarks>
        public static void TOC()
        {

            string pathFile = @"..\..\..\toc.docx";
            string resultFile = "UpdatedTOC.docx";

            // Load a .docx document with TOC.
            DocumentCore dc = DocumentCore.Load(pathFile);

            Paragraph p = new Paragraph(dc);
            p.Content.Start.Insert("I was born in the year 1632, in the city of York, of a good family, though not of that country, " +
                "my father being a foreigner of Bremen, who settled first at Hull.  He got a good estate by merchandise, and leaving " +
                "off his trade, lived afterwards at York, from whence he had married my mother, whose relations were named Robinson, " +
                "a very good family in that country, and from whom I was called Robinson Kreutznaer; but, by the usual corruption " +
                "of words in England, we are now called-nay we call ourselves and write our name-Crusoe; and so my companions always " +
                "called me. I had two elder brothers, one of whom was lieutenant-colonel to an English regiment of foot in Flanders, " +
                "formerly commanded by the famous Colonel Lockhart, and was killed at the battle near Dunkirk against the Spaniards.",

                new CharacterFormat() { Size = 28 });
            p.ParagraphFormat.Alignment = HorizontalAlignment.Justify;

            // Insert the paragraph as 6th element in the 1st section.
            dc.Sections[0].Blocks.Insert(5, p);

            Paragraph p1 = new Paragraph(dc);
            p1.Content.Start.Insert("That evil influence which carried me first away from my father’s house-which hurried me into the " +
                "wild and indigested notion of raising my fortune, and that impressed those conceits so forcibly upon me as to make me " +
                "deaf to all good advice, and to the entreaties and even the commands of my father-I say, the same influence, whatever " +
                "it was, presented the most unfortunate of all enterprises to my view; and I went on board a vessel bound to the coast " +
                "of Africa; or, as our sailors vulgarly called it, a voyage to Guinea. It was my great misfortune that in all these " +
                "adventures I did not ship myself as a sailor; when, though I might indeed have worked a little harder than ordinary, " +
                "yet at the same time I should have learnt the duty and office of a fore-mast man, and in time might have qualified " +
                "myself for a mate or lieutenant, if not for a master.  But as it was always my fate to choose for the worse, so I did " +
                "here; for having money in my pocket and good clothes upon my back, I would always go on board in the habit of " +
                "a gentleman; and so I neither had any business in the ship, nor learned to do any.",

                new CharacterFormat() { Size = 28 });
            p1.ParagraphFormat.Alignment = HorizontalAlignment.Justify;

            // Insert the paragraph as 10th element in the 1st section.
            dc.Sections[0].Blocks.Insert(9, p1);

            // Update TOC (TOC can be updated only after all document content is added).
            TableOfEntries toc = (TableOfEntries)dc.GetChildElements(true, ElementType.TableOfEntries).FirstOrDefault();

            toc.Update();

            // Update TOC's page numbers.
            // Page numbers are automatically updated in that case.
            dc.GetPaginator(new PaginatorOptions() { UpdateFields = true });

            // Change default character formatting for all text inside TOC
            CharacterFormat cf = new CharacterFormat();
            cf.Size = 20;
            cf.FontColor = Color.Blue;
            foreach (Inline inline in toc.GetChildElements(true, ElementType.Run, ElementType.SpecialCharacter))
            {
                if (inline is Run)
                    ((Run)inline).CharacterFormat = cf.Clone();
                else
                    ((SpecialCharacter)inline).CharacterFormat = cf.Clone();
            }

            // Save the document as new DOCX file.
            dc.Save(resultFile);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultFile) { UseShellExecute = true });
        }
    }
}
