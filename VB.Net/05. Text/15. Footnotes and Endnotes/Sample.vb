Imports System
Imports SautinSoft.Document
Imports SautinSoft.Document.Tables

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			CreateFootnoteAndEndnote()
		End Sub
		''' <summary>
		''' Creates a document with a footnote and endnote.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/text-footnotes.php
		''' </remarks>
		Private Shared Sub CreateFootnoteAndEndnote()
			Dim filePath As String = "FootnoteAndEndnote.docx"

			Dim dc As New DocumentCore()

			Dim p1 As Paragraph = New Paragraph(dc, New Run(dc, "Every evening the young Fisherman went out upon the sea, and " & "threw his nets into the water.") With {
				.CharacterFormat = New CharacterFormat() With {.Size = 24}
			})

			Dim p2 As Paragraph = New Paragraph(dc, New Run(dc, "When the wind blew from the land he caught nothing, or but little at best," & " for it was a bitter and black-winged wind, and rough waves rose up to meet it") With {
				.CharacterFormat = New CharacterFormat() With {.Size = 24}
			},
			New Note(dc, NoteType.Endnote, "This is the endnote.") With {
				.CharacterFormat = New CharacterFormat() With {
					.Size = 36,
					.Superscript = True
				}
			},
			New Run(dc, " . But when the wind blew to the shore, the " & "fish came in from the deep, and swam into the meshes of his nets, " & "and he took them to the market-place and sold them.") With {
				.CharacterFormat = New CharacterFormat() With {.Size = 24}
			},
			New Note(dc, NoteType.Endnote, "This is the endnote.") With {
				.CharacterFormat = New CharacterFormat() With {
					.Size = 36,
					.Superscript = True
				}
			})

			Dim p3 As Paragraph = New Paragraph(dc, New Run(dc, "Every evening he went out upon the sea, and one evening the net " & " was so heavy that hardly could he draw it into the boat") With {
				.CharacterFormat = New CharacterFormat() With {.Size = 24}
			},
			New Note(dc, NoteType.Footnote, "This is the footnote.") With {
				.CharacterFormat = New CharacterFormat() With {
					.Size = 36,
					.Superscript = True
				}
			},
			New Run(dc, ". And he laughed, and said to himself, ‘Surely I  " & "have caught all the fish that swim, or snared some dull monster that will be a marvel to men, or some thing  " & "of horror that the great Queen will desire,’ and putting forth all his strength, he tugged at the coarse ropes  " & "till, like lines of blue enamel round a vase of bronze, the long veins rose up on his arms.") With {
				.CharacterFormat = New CharacterFormat() With {.Size = 24}
			},
			New Note(dc, NoteType.Endnote, "This is the endnote.") With {
				.CharacterFormat = New CharacterFormat() With {
					.Size = 36,
					.Superscript = True
				}
			},
			New Run(dc, "He tugged at the  " & "thin ropes, and nearer and nearer came the circle of flat corks, and the net rose at last to the top of the water.") With {
				.CharacterFormat = New CharacterFormat() With {.Size = 24}
			})

			Dim footnote As New Note(dc, NoteType.Footnote, "This is the footnote.")
			footnote.CharacterFormat.Superscript = True
			footnote.CharacterFormat.Size = 36

			p1.Content.End.Insert(footnote.Content)
			p3.Content.End.Insert(footnote.Content)
			dc.Content.End.Insert(p1.Content)
			dc.Content.End.Insert(p2.Content)
			dc.Content.End.Insert(p3.Content)

			dc.Save(filePath)

			' Show the result.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
