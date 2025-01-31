Option Infer On
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports SautinSoft.Document.Tables

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			' Get your free trial key here:   
			' https://sautinsoft.com/start-for-free/
			DisplayTable()
		End Sub

		''' <summary>
		''' How to display a table always at the bottom of the page.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-display-table-always-at-the-bottom-of-the-page-in-pdf-net-csharp-vb.php
		''' </remarks>
		Private Shared Sub DisplayTable()
			' Is there a way to display a table always at the bottom of the page?
			' Yes, place the table in the document footer.

			' Create 5-pages document and place an unique table
			' at the bottom of each page.
			Dim dc = New DocumentCore()

			Dim resultPath As String = "Result.pdf"
			Dim pagesText() As String = {"March", "April", "May", "June", "July"}

			For page As Integer = 0 To pagesText.Length - 1
				Dim s = New Section(dc)
				dc.Sections.Add(s)
				' Write some text content
				Dim p = New Paragraph(dc)
				p.ParagraphFormat.Alignment = HorizontalAlignment.Center
				Dim cf = New CharacterFormat() With {.Size = 100.0F
				}
				p.Inlines.Add(New Run(dc, pagesText(page), cf))
				s.Blocks.Add(p)

				' Place the table in the document footer.                
				AddTableToFooter(dc, s)
			Next page

			dc.Save(resultPath)
			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
		End Sub
		Private Shared Sub AddTableToFooter(ByVal dc As DocumentCore, ByVal s As Section)
			Dim rand As New Random()
			Dim p As Paragraph

			Dim tableText() As String = {"Item", "Quantity", "Price"}
			Dim t As New Table(dc)
			t.TableFormat.PreferredWidth = New TableWidth(100, TableWidthUnit.Percentage)
			t.TableFormat.Alignment = HorizontalAlignment.Center

			' Table header
			Dim rowHdr = New TableRow(dc)
			For Each cellText In tableText
				Dim cellHdr = New TableCell(dc)
				cellHdr.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 0.5F)
				cellHdr.CellFormat.BackgroundColor = New Color(rand.Next(Int32.MaxValue))
				cellHdr.CellFormat.PreferredWidth = New TableWidth(100.0F / 3.0F, TableWidthUnit.Percentage)
				p = New Paragraph(dc, cellText)
				p.ParagraphFormat.Alignment = HorizontalAlignment.Center
				cellHdr.Blocks.Add(p)
				rowHdr.Cells.Add(cellHdr)
			Next cellText
			t.Rows.Add(rowHdr)
			' Table body
			Dim rowCount As Integer = rand.Next(1, 20)
			For r As Integer = 0 To rowCount - 1
				Dim row = New TableRow(dc)
				For Each cellText In tableText
					Dim cell = New TableCell(dc)
					cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 0.5F)
					cell.CellFormat.BackgroundColor = Color.White
					cell.CellFormat.PreferredWidth = New TableWidth(100.0F / 3.0F, TableWidthUnit.Percentage)
					p = New Paragraph(dc, rand.Next(100).ToString())
					p.ParagraphFormat.Alignment = HorizontalAlignment.Center
					cell.Blocks.Add(p)
					row.Cells.Add(cell)
				Next cellText
				t.Rows.Add(row)
			Next r
			' Move table to page footer
			Dim footer As New HeaderFooter(dc, HeaderFooterType.FooterDefault)
			s.HeadersFooters.Add(footer)
			footer.Blocks.Add(t)
		End Sub
	End Class
End Namespace