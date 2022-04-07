Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports SautinSoft.Document.Tables
Imports System.Linq

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			TableRowFormatting()
		End Sub
		''' <summary>
		''' Shows how to set a height for a table row, repeat a row as header on each page, shift a row by N columns to the right.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/tablerow-format.php
		''' </remarks>
		Private Shared Sub TableRowFormatting()
			Dim docxPath As String = "FormattedTable.docx"

			' Let's create document.
			Dim dc As New DocumentCore()

			Dim s As New Section(dc)
			dc.Sections.Add(s)


			Dim rows As Integer = 30
			Dim columns As Integer = 5

			Dim t As New Table(dc, rows, columns)
			t.TableFormat.PreferredWidth = New TableWidth(100, TableWidthUnit.Percentage)
			t.TableFormat.Borders.SetBorders(MultipleBorderTypes.All, BorderStyle.Single, Color.DarkGray, 1)
			t.TableFormat.AutomaticallyResizeToFitContents = False
			s.Blocks.Add(t)

			' Specify row height:
			' 10 mm - for odd rows.
			' 15 mm - for even rows.
			Dim oddHeight As Double = LengthUnitConverter.Convert(10, LengthUnit.Millimeter, LengthUnit.Point)
			Dim evenHeight As Double = LengthUnitConverter.Convert(15, LengthUnit.Millimeter, LengthUnit.Point)
			For r As Integer = 0 To t.Rows.Count - 1
				Dim row As TableRow = t.Rows(r)
				If r Mod 2 <> 0 Then
					row.RowFormat.Height = New TableRowHeight(evenHeight, HeightRule.AtLeast)
				Else
					row.RowFormat.Height = New TableRowHeight(oddHeight, HeightRule.AtLeast)
				End If
			Next r

			' Add the table caption - mark the specific row (for example: 0) to repeat on each page.
			Dim firstRow As TableRow = t.Rows(0)
			' Repeate as header row at the top of each page.
			' Note: Only the first row in the table can be set up as header.
			firstRow.RowFormat.RepeatOnEachPage = True

			' Merge all cells into a one in the first row (Caption).
			Dim colSpan As Integer = firstRow.Cells.Count
			For c As Integer = firstRow.Cells.Count - 1 To 1 Step -1
				firstRow.Cells.RemoveAt(c)
			Next c
			' Specify how many columns this cell will take up.
			firstRow.Cells(0).ColumnSpan = colSpan

			' Set the table caption in the first row and first cell. 
			Dim p As New Paragraph(dc)
			p.Inlines.Add(New Run(dc, "This is the Row 0 (RepeatOnEachPage = true)", New CharacterFormat() With {
				.FontColor = Color.Blue,
				.Size = 20
			}))
			p.ParagraphFormat.Alignment = HorizontalAlignment.Center
			t.Rows(0).Cells(0).Blocks.Add(p)

			' Another interesting properties of TableRowFormat:
			' GridBefore and GridAfter
			' Add "Total" at the end of the table.
			Dim rowTotal As New TableRow(dc)
			rowTotal.Cells.Add(New TableCell(dc))
			rowTotal.Cells(0).Content.Start.Insert(String.Format("Total rows: {0}", rows), New CharacterFormat() With {
				.FontColor = Color.Red,
				.Size = 30
			})

			' Shift the rowTotal to the right corner.
			' In our case, shift on 4 columns.
			rowTotal.RowFormat.GridBefore = columns-1


			rowTotal.RowFormat.Height = New TableRowHeight(evenHeight, HeightRule.AtLeast)
			t.Rows.Add(rowTotal)

			' Save our document into DOCX format.
			dc.Save(docxPath)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(docxPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
