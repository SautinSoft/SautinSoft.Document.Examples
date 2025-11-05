Option Infer On

Imports SautinSoft.Document
Imports SautinSoft.Document.Tables
Imports System
Imports System.IO
Imports System.Linq
Imports System.Reflection.Metadata

Namespace Sample
	Friend Class Sample
		Shared Sub Main(ByVal args() As String)
			' Get your free trial key here:   
			' https://sautinsoft.com/start-for-free/

			RTLTable()
		End Sub

		''' <summary>
		''' How to add Rigth-to-Left text in a table.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/right-to-left-table.php
		''' </remarks>
		Public Shared Sub RTLTable()
			Dim sourcePath As String = "..\..\..\RTL.docx"
			Dim destPath As String = "RTL.pdf"
			Dim dc As DocumentCore = DocumentCore.Load(sourcePath)
			' Show line numbers on the right side of the page
			Dim pageSetup = dc.Sections(0).PageSetup
			pageSetup.LineNumberRestartSetting = LineNumberRestartSetting.Continuous

			' Create a new right-to-left paragraph
			Dim paragraph As New Paragraph(dc)
			paragraph.ParagraphFormat.RightToLeft = True
			paragraph.Inlines.Add(New Run(dc, "أخذ عن موالية الإمتعاض"))

			dc.Sections(0).Blocks.Add(paragraph)

			' Create a right-to-left table with some data inside. 
			Dim table As New Table(dc)

			table.TableFormat.PreferredWidth = New TableWidth(100, TableWidthUnit.Percentage)
			Dim row = New TableRow(dc)
			table.TableFormat.Borders.SetBorders(MultipleBorderTypes.Inside Or MultipleBorderTypes.Outside, BorderStyle.Single, Color.Black, 1)
			table.Rows.Add(row)

			Dim firstCellPara = New Paragraph(dc, "של תיבת תרומה מלא")
			firstCellPara.ParagraphFormat.RightToLeft = True
			row.Cells.Add(New TableCell(dc, firstCellPara))

			Dim secondCellPara = New Paragraph(dc, "200")
			row.Cells.Add(New TableCell(dc, secondCellPara))
			dc.Sections(0).Blocks.Add(table)

			' Save the document as PDF.
			dc.Save(destPath, New PdfSaveOptions())

			' Show the source and the dest documents.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(sourcePath) With {.UseShellExecute = True})
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(destPath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
