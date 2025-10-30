using System.IO;
using SautinSoft.Document;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free trial key here:   
            // https://sautinsoft.com/start-for-free/

            TransformationInvisibleText();     
        }
        /// <summary>
        /// Transformation Invisible Text (PDF to PDF/DOCX, DOCX to PDF/DOCX).
        /// </summary>
		/// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/transformation_invisible_text.php
        /// </remarks>
        static void TransformationInvisibleText()
        {
            string loadPath = @"..\..\..\invisible_text.pdf";
            string savePathPDF = @"..\..\..\visible_text.pdf";
            string savePathDocx = @"..\..\..\visible_text.docx";

            //DocumentCore.Serial = "put your serial here";
            DocumentCore pdf = DocumentCore.Load(loadPath, new PdfLoadOptions()
            {
                 PreserveImages = false,
                 ShowInvisibleText = true,
                 PreserveEmbeddedFonts = true
            });

            CharacterFormat singleCharFormat = new CharacterFormat()
            {            
                FontColor = Color.Black
            };

            pdf.DefaultCharacterFormat = singleCharFormat.Clone();

            foreach (Element element in pdf.GetChildElements(true, ElementType.Run, ElementType.Paragraph))
            {
                if (element is Run)
                   (element as Run).CharacterFormat = singleCharFormat.Clone();
                else if (element is Paragraph)
                    (element as Paragraph).CharacterFormatForParagraphMark = singleCharFormat.Clone();
            }
            pdf.Save(savePathPDF);
            pdf.Save(savePathDocx);
        }
        }
    }
