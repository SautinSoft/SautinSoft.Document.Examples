Imports System
Imports SautinSoft.Document
Imports SautinSoft.Document.Tables
Imports System.Collections.Generic

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			InsertingTable()
		End Sub
		''' <summary>
		''' Creates a table using DocumentBuilder and saves it in a desired format.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/documentbuilder-inserting-table.php
		''' </remarks>

		Private Shared Sub InsertingTable()
			Dim dc As New DocumentCore()
			Dim db As New DocumentBuilder(dc)

			' Create a new table with preferred width.
			Dim table As Table = db.StartTable()
			db.TableFormat.PreferredWidth = New TableWidth(LengthUnitConverter.Convert(5, LengthUnit.Inch, LengthUnit.Point), TableWidthUnit.Point)

			' Specify formatting of cells and alignment.
			db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Single, Color.Green, 1)
			db.CellFormat.VerticalAlignment = VerticalAlignment.Top
			db.ParagraphFormat.Alignment = HorizontalAlignment.Center

			' Specify height of rows and write text.
			db.RowFormat.Height = New TableRowHeight(105.0F, HeightRule.Exact)
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
			db.ParagraphFormat.Alignment = HorizontalAlignment.Left

			' Specify height of rows and write text.
			db.RowFormat.Height = New TableRowHeight(150.0F, HeightRule.Exact)
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
			db.ParagraphFormat.Alignment = HorizontalAlignment.Right

			' Specify height of rows and write text
			db.RowFormat.Height = New TableRowHeight(125.0F, HeightRule.Exact)
			db.InsertCell()
			db.Write("This is Row 3 Cell 1")
			db.InsertCell()
			db.Write("This is Row 3 Cell 2")
			db.InsertCell()
			db.Write("This is Row 3 Cell 3")
			db.EndRow()
			db.EndTable()

			' Save our document into DOCX format.
			Dim filePath As String = "Result.docx"
			dc.Save(filePath)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace