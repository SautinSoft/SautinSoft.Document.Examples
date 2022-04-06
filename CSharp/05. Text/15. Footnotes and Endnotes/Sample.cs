using System;
using SautinSoft.Document;
using SautinSoft.Document.Tables;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            CreateFootnoteAndEndnote();
        }
        /// <summary>
        /// Creates a document with a footnote and endnote.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/text-footnotes.php
        /// </remarks>
        static void CreateFootnoteAndEndnote()
        {
            string filePath = @"FootnoteAndEndnote.docx";

            DocumentCore dc = new DocumentCore();

            Paragraph p1 = new Paragraph(dc, new Run(dc, "Every evening the young Fisherman went out upon the sea, and " +
                "threw his nets into the water.")
            {
                CharacterFormat = new CharacterFormat()
                {
                    Size = 24,
                }
            });

            Paragraph p2 = new Paragraph(dc,
                           new Run(dc, "When the wind blew from the land he caught nothing, or but little at best," +
                           " for it was a bitter and black-winged wind, and rough waves rose up to meet it")
                           {
                               CharacterFormat = new CharacterFormat()
                               {
                                   Size = 24,
                               }
                           },
                           new Note(dc, NoteType.Endnote, "This is the endnote.")
                           {
                               CharacterFormat = new CharacterFormat()
                               {
                                   Size = 36,
                                   Superscript = true
                               }
                           },
                           new Run(dc, " . But when the wind blew to the shore, the " +
                           "fish came in from the deep, and swam into the meshes of his nets, " +
                            "and he took them to the market-place and sold them.")
                           {
                               CharacterFormat = new CharacterFormat()
                               {
                                   Size = 24
                               }
                           },
                           new Note(dc, NoteType.Endnote, "This is the endnote.")
                           {
                               CharacterFormat = new CharacterFormat()
                               {
                                   Size = 36,
                                   Superscript = true
                               }
                           });

            Paragraph p3 = new Paragraph(dc,
                new Run(dc, "Every evening he went out upon the sea, and one evening the net " +
                " was so heavy that hardly could he draw it into the boat")
                {
                    CharacterFormat = new CharacterFormat()
                    {
                        Size = 24,
                    }
                },
                new Note(dc, NoteType.Footnote, "This is the footnote.")
                {
                    CharacterFormat = new CharacterFormat()
                    {
                        Size = 36,
                        Superscript = true
                    }
                },
                new Run(dc, ". And he laughed, and said to himself, ‘Surely I  " +
                "have caught all the fish that swim, or snared some dull monster that will be a marvel to men, or some thing  " +
                "of horror that the great Queen will desire,’ and putting forth all his strength, he tugged at the coarse ropes  " +
                "till, like lines of blue enamel round a vase of bronze, the long veins rose up on his arms.")
                {
                    CharacterFormat = new CharacterFormat()
                    {
                        Size = 24,
                    }
                },
                new Note(dc, NoteType.Endnote, "This is the endnote.")
                {
                    CharacterFormat = new CharacterFormat()
                    {
                        Size = 36,
                        Superscript = true
                    }
                },
                new Run(dc, "He tugged at the  " +
                "thin ropes, and nearer and nearer came the circle of flat corks, and the net rose at last to the top of the water.")
                {
                    CharacterFormat = new CharacterFormat()
                    {
                        Size = 24,
                    }
                });

            Note footnote = new Note(dc, NoteType.Footnote, "This is the footnote.");
            footnote.CharacterFormat.Superscript = true;
            footnote.CharacterFormat.Size = 36;

            p1.Content.End.Insert(footnote.Content);
            p3.Content.End.Insert(footnote.Content);
            dc.Content.End.Insert(p1.Content);
            dc.Content.End.Insert(p2.Content);
            dc.Content.End.Insert(p3.Content);

            dc.Save(filePath);

            // Show the result.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}