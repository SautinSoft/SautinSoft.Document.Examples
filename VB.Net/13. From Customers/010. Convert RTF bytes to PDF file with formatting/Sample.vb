Imports System
Imports System.IO
Imports SautinSoft.Document
Imports SautinSoft.Document.Tables
Imports System.Linq
Imports System.Text.RegularExpressions

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			ConvertRtfBytesToPdfFile()
		End Sub
		''' <summary>
		''' Convert RTF bytes to PDF file with formatting.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-convert-rtf-bytes-to-pdf-with-formatting-in-csharp-vb-net.php
		''' </remarks>
		Public Shared Sub ConvertRtfBytesToPdfFile()
			' Get document bytes.
			Dim fileBytes() As Byte = File.ReadAllBytes("..\..\..\example.rtf")

			Dim PdfPath As String = "result.pdf"

			Dim dc As DocumentCore = Nothing
			Dim regex As New Regex("formatting", RegexOptions.IgnoreCase)

			' Create a MemoryStream.
			Using ms As New MemoryStream(fileBytes)
				' Load a document from the MemoryStream.
				' Specifying RtfLoadOptions we explicitly set that a loadable document is RTF.
				dc = DocumentCore.Load(ms, New RtfLoadOptions())
			End Using

			' Add a new section in the document.
			Dim s As New Section(dc)
			dc.Sections.Add(s)

			' Create a new table with two rows and three columns inside.
			Dim table As New Table(dc)
			Dim width As Double = LengthUnitConverter.Convert(100, LengthUnit.Millimeter, LengthUnit.Point)
			table.TableFormat.PreferredWidth = New TableWidth(width, TableWidthUnit.Point)
			table.TableFormat.Alignment = HorizontalAlignment.Center

			Dim counter As Integer = 0

			' Add rows.
			Dim rows As Integer = 2
			Dim columns As Integer = 3
			For r As Integer = 0 To rows - 1
				Dim row As New TableRow(dc)

				' Add columns.
				For c As Integer = 0 To columns - 1
					Dim cell As New TableCell(dc)

					' Set cell formatting and width.
					cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dashed, Color.Brown, 2.0)

					' Set the same width for each column.
					cell.CellFormat.PreferredWidth = New TableWidth(width / columns, TableWidthUnit.Point)

					If counter Mod 2 = 1 Then
						cell.CellFormat.BackgroundColor = New Color("#FF0000")
					End If

					row.Cells.Add(cell)

					' Let's add a paragraph with text into the each column.
					Dim pa As New Paragraph(dc)
					pa.ParagraphFormat.Alignment = HorizontalAlignment.Center
					pa.ParagraphFormat.SpaceBefore = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point)
					pa.ParagraphFormat.SpaceAfter = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point)

					pa.Content.Start.Insert(String.Format("{0}", ChrW(counter + AscW("A"c))), New CharacterFormat() With {
						.FontName = "Arial",
						.FontColor = New Color("#000000"),
						.Size = 12.0
					})
					cell.Blocks.Add(pa)
					counter += 1
				Next c
				table.Rows.Add(row)

			Next r
				' Create a new header with formatted text.
				Dim header As New HeaderFooter(dc, HeaderFooterType.HeaderDefault)
				header.Content.Start.Insert(table.Content)
				For Each s1 As Section In dc.Sections
					s1.HeadersFooters.Add(header.Clone(True))
				Next s1

				' Add the header into HeadersFooters collection of the 1st section.
				's1.HeadersFooters.Add(header);

				' Create a new footer with formatted text.
				Dim footer As New HeaderFooter(dc, HeaderFooterType.FooterDefault)
				footer.Content.Start.Insert(table.Content)
				For Each s1 As Section In dc.Sections
					s1.HeadersFooters.Add(footer.Clone(True))
				Next s1

				' Add the footer into HeadersFooters collection of the 1st section.
				's1.HeadersFooters.Add(footer);

				For Each item As ContentRange In dc.Content.Find(regex).Reverse()
				' Replace all text "formatting" on "FORMATTING!!!".
				item.Replace("FORMATTING!!!", New CharacterFormat() With {
					.BackgroundColor = Color.Yellow,
					.FontName = "Arial",
					.Size = 16.0
				})
				Next item

				' Save our result as a PDF file.
			dc.Save(PdfPath)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(PdfPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace