Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports SautinSoft.Document.Tables
Imports System


Namespace Sample
	Friend Class Sample

		Shared Sub Main(ByVal args() As String)
			TableInHeader()
		End Sub

		''' <summary>
		''' How to add a table in header or footer into a document.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/headersfooters.php
		''' </remarks>
		Public Shared Sub TableInHeader()
			Dim documentPath As String = "TableHeader.docx"

			' Let's create a simple document.
			Dim dc As New DocumentCore()

			' Add a new section in the document.
			Dim s As New Section(dc)
			dc.Sections.Add(s)

			' Create a new table with two rows and three columns inside
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
						cell.CellFormat.BackgroundColor = Color.Red
					End If

					row.Cells.Add(cell)

					' Let's add a paragraph with text into the each column.
					Dim pa As New Paragraph(dc)
					pa.ParagraphFormat.Alignment = HorizontalAlignment.Center
					pa.ParagraphFormat.SpaceBefore = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point)
					pa.ParagraphFormat.SpaceAfter = LengthUnitConverter.Convert(3, LengthUnit.Millimeter, LengthUnit.Point)

					pa.Content.Start.Insert(String.Format("{0}", ChrW(counter + AscW("A"c))), New CharacterFormat() With {
						.FontName = "Arial",
						.FontColor = Color.Black,
						.Size = 14.0
					})
					cell.Blocks.Add(pa)
					counter += 1
				Next c
				table.Rows.Add(row)
			Next r

			' Create a new header with formatted text.
			Dim header As New HeaderFooter(dc, HeaderFooterType.HeaderDefault)
			header.Content.Start.Insert(table.Content)

			' Add the header into HeadersFooters collection of the 1st section.
			s.HeadersFooters.Add(header)

			' Create a new footer with formatted text.
			Dim footer As New HeaderFooter(dc, HeaderFooterType.FooterDefault)
			footer.Content.Start.Insert(table.Content)

			' Add the footer into HeadersFooters collection of the 1st section.
			s.HeadersFooters.Add(footer)

			' Save the document into DOCX format.
			dc.Save(documentPath)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
