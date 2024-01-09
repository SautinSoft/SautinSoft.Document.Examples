Option Infer On
Imports System
Imports System.IO
Imports System.Collections.Generic
Imports SautinSoft.Document
Imports System.Linq
Imports SautinSoft.Document.Tables

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			TableWithDocumentCore()
			TableWithDocumentBuilder()
		End Sub
                ''' Get your free 30-day key here:   
                ''' https://sautinsoft.com/start-for-free/
		''' <summary>
		''' This sample shows how to creating a custom table with formatting using DocumentCore or DocumentBuilder classes.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/from-customers-creating-custom-table-with-formatting-in-csharp-vb-net.php
		''' </remarks>
		Public Shared Sub TableWithDocumentCore()
			Dim documentPath As String = "tableDC.pdf"

			' Let's create a new document.
			Dim dc As New DocumentCore()

			Dim s As New Section(dc)

			Dim tf As New TableFormat()
			tf.Borders.ClearBorders()
			tf.AutomaticallyResizeToFitContents = False

			Dim table = New Table(dc)

			' Add columns with specified width.
			table.Columns.Add(New TableColumn(60))
			table.Columns.Add(New TableColumn(120))
			table.Columns.Add(New TableColumn(180))

			' Add rows with specified height.
			table.Rows.Add(New TableRow(dc))
			table.Rows(0).RowFormat.Height = New TableRowHeight(30, HeightRule.AtLeast)
			table.Rows.Add(New TableRow(dc))
			table.Rows(1).RowFormat.Height = New TableRowHeight(60, HeightRule.AtLeast)
			table.Rows.Add(New TableRow(dc))
			table.Rows(2).RowFormat.Height = New TableRowHeight(90, HeightRule.AtLeast)

			For r As Integer = 0 To 2
				For c As Integer = 0 To 2
					' Add cell.
					Dim cell = New TableCell(dc)
					table.Rows(r).Cells.Add(cell)

					' Set cell's vertical alignment.
					cell.CellFormat.VerticalAlignment = CType(r, VerticalAlignment)

					' Add cell content.
					Dim paragraph = New Paragraph(dc, $"Cell ({r + 1},{c + 1})")
					cell.Blocks.Add(paragraph)

					' Set cell content's horizontal alignment.
					paragraph.ParagraphFormat.Alignment = CType(c, HorizontalAlignment)

					If (r + c) Mod 2 = 0 Then
						' Set cell's background and borders.
						cell.CellFormat.BackgroundColor = New Color(255, 242, 204)
						cell.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Double, Color.Red, 1)
					End If
				Next c
			Next r

			dc.Sections.Add(New Section(dc, table))

			' Save our document into PDF format.
			dc.Save(documentPath, New PdfSaveOptions() With {.Compliance = PdfCompliance.PDF_A1a})

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(documentPath) With {.UseShellExecute = True})
		End Sub

		Public Shared Sub TableWithDocumentBuilder()
			Dim dc As New DocumentCore()
			Dim db As New DocumentBuilder(dc)

			' Create a new table with preferred width.
			Dim table As Table = db.StartTable()

		   ' db.TableFormat.PreferredWidth = new TableWidth(LengthUnitConverter.Convert(5, LengthUnit.Inch, LengthUnit.Point), TableWidthUnit.Point);
			db.TableFormat.AutomaticallyResizeToFitContents = False

			' Specify formatting of cells and alignment.
			db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Green, 1)
			db.CellFormat.VerticalAlignment = VerticalAlignment.Top
			table.Columns.Add(New TableColumn() With {.PreferredWidth = 100})
			db.ParagraphFormat.Alignment = HorizontalAlignment.Center

			' Specify height of rows and write text.
			db.RowFormat.Height = New TableRowHeight(105F, HeightRule.Exact)

			db.InsertCell()
			db.Write("This is Row 1 Cell 1")
			db.InsertCell()
			db.Write("This is Row 1 Cell 2")
			db.InsertCell()
			db.Write("This is Row 1 Cell 3")
			db.EndRow()

			' Specify formatting of cells and alignment.
			db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 1)
			db.CellFormat.VerticalAlignment = VerticalAlignment.Center
			table.Columns.Add(New TableColumn() With {.PreferredWidth = 100})
			db.ParagraphFormat.Alignment = HorizontalAlignment.Left

			' Specify height of rows and write text.
			db.RowFormat.Height = New TableRowHeight(150F, HeightRule.Exact)
			db.InsertCell()
			db.Write("This is Row 2 Cell 1")
			db.InsertCell()
			db.Write("This is Row 2 Cell 2")
			db.InsertCell()
			db.Write("This is Row 2 Cell 3")
			db.EndRow()

			' Specify formatting of cells and alignment.
			db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Orange, 1)
			db.CellFormat.VerticalAlignment = VerticalAlignment.Bottom
			table.Columns.Add(New TableColumn() With {.PreferredWidth = 150})
			db.ParagraphFormat.Alignment = HorizontalAlignment.Right

			' Specify height of rows and write text
			db.RowFormat.Height = New TableRowHeight(125F, HeightRule.Exact)
			db.InsertCell()
			db.Write("This is Row 3 Cell 1")
			db.InsertCell()
			db.Write("This is Row 3 Cell 2")
			db.InsertCell()
			db.Write("This is Row 3 Cell 3")
			db.EndRow()
			db.EndTable()

			' Save our document into DOCX format.
			Dim filePath As String = "tableDB.docx"
			dc.Save(filePath)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace