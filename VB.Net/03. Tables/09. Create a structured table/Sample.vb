Option Infer On

Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.IO
Imports System.Linq
Imports System.Xml.Linq
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing
Imports SautinSoft.Document.Tables

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
				StructuredTable()
		End Sub
                ''' Get your free trial key here:   
                ''' https://sautinsoft.com/start-for-free/
		''' <summary>
		''' Creating a structured table with data of different formats.
		''' </summary>
		''' <remarks>
		''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-structured-table.php
		''' </remarks>
		Private Shared Sub StructuredTable()
			Dim documentCore As New DocumentCore()
			Dim imageData() As Byte = File.ReadAllBytes("../../../image/smile.jpg")
			Dim TableHeaderStyle As ParagraphStyle = CType(Style.CreateStyle(StyleTemplateType.Normal, documentCore), ParagraphStyle)
			TableHeaderStyle.Name = "TableHeaderStyle"
			TableHeaderStyle.ParagraphFormat.Alignment = HorizontalAlignment.Left
			TableHeaderStyle.ParagraphFormat.LeftIndentation = 6
			TableHeaderStyle.CharacterFormat.FontColor = SautinSoft.Document.Color.White
			TableHeaderStyle.CharacterFormat.FontName = "Segoe UI"
			TableHeaderStyle.CharacterFormat.Size = 9
			documentCore.Styles.Add(TableHeaderStyle)

			Dim TableContentStyle_1 As ParagraphStyle = CType(Style.CreateStyle(StyleTemplateType.Normal, documentCore), ParagraphStyle)
			TableContentStyle_1.Name = "TableContentStyle_1"
			TableContentStyle_1.ParagraphFormat.Alignment = HorizontalAlignment.Left
			TableContentStyle_1.ParagraphFormat.LeftIndentation = 6
			'TableContentStyle_1.CharacterFormat.FontName = "Segoe UI";
			TableContentStyle_1.CharacterFormat.Size = 9
			documentCore.Styles.Add(TableContentStyle_1)

			Dim TableOutlineStyle As TableStyle = CType(Style.CreateStyle(StyleTemplateType.TableNormal, documentCore), TableStyle)
			TableOutlineStyle.Name = "TableOutlineStyle"
			TableOutlineStyle.ParagraphFormat.SpaceAfter = 10
			documentCore.Styles.Add(TableOutlineStyle)

			Dim table As New Table(documentCore)
			table.TableFormat.PreferredWidth = New TableWidth(100, TableWidthUnit.Percentage)
			table.TableFormat.AutomaticallyResizeToFitContents = True
			table.TableFormat.Alignment = HorizontalAlignment.Left
			table.TableFormat.Style = TableOutlineStyle

			Dim headerRow As New TableRow(documentCore)
			If True Then
				Dim cell1 As New TableCell(documentCore)
				cell1.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Thick, SautinSoft.Document.Color.White, 1.5)
				cell1.CellFormat.BackgroundColor = New SautinSoft.Document.Color("#0054A6")
				cell1.CellFormat.Padding = New Padding(0, 3, 0, 0)
				Dim pHeader As New Paragraph(documentCore, "HeaderName")
				pHeader.ParagraphFormat.Style = TableHeaderStyle
				cell1.Blocks.Add(pHeader)
				headerRow.Cells.Add(cell1)
			End If
			If True Then
				Dim cell2 As New TableCell(documentCore)
				cell2.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Thick, SautinSoft.Document.Color.White, 1.5)
				cell2.CellFormat.BackgroundColor = New SautinSoft.Document.Color("#0054A6")
				cell2.CellFormat.Padding = New Padding(0, 3, 0, 0)
				Dim pHeaderContent As New Paragraph(documentCore, "HeaderContent")
				pHeaderContent.ParagraphFormat.Style = TableHeaderStyle
				cell2.Blocks.Add(pHeaderContent)
				headerRow.Cells.Add(cell2)
			End If
			headerRow.RowFormat.Height = New TableRowHeight(15, HeightRule.Auto)
			headerRow.RowFormat.RepeatOnEachPage = True
			table.Rows.Add(headerRow)


			Dim x As String = "Detects if your application experiences an abnormal rise in the rate of HTTP requests or dependency calls that are reported as failed." &
				" The anomaly detection uses machine learning algorithms and occurs in near real time, " &
				"therefore there's no need to define a frequency for this signal.<br><br>To help you triage and diagnose the problem, an analysis of the characteristics of the failures and related telemetry is provided with the detection. " &
				"This feature works for any app, hosted in the cloud or on your own servers, that generates request or dependency telemetry - for example, " &
				"if you have a worker role that calls <a class=""ext-smartDetecor-link"" href=""https://docs.microsoft.com/azure/application-insights/app-insights-api-custom-events-metrics#trackrequest"" target=""_blank"">TrackRequest()</a> " &
				"or <a class=""ext-smartDetecor-link"" href=""https://docs.microsoft.com/azure/application-insights/app-insights-api-custom-events-metrics#trackdependency"" target=""_blank"">TrackDependency()</a>." &
				"<br/><br/><a class=""ext-smartDetecor-link"" href=""https://docs.microsoft.com/azure/azure-monitor/app/proactive-failure-diagnostics"" target=""_blank"">Learn more about Failure Anomalies</a><br><br>" &
				"<p style=""font-size: 13px; font-weight: 700;"">A note about your data privacy:</p><br><br>The service is entirely automatic and only you can see these notifications. " &
				"<a class=""ext-smartDetecor-link"" href=""https://docs.microsoft.com/en-us/azure/azure-monitor/app/data-retention-privacy"" target=""_blank"">Read more about data privacy</a><br><br>Smart Alerts conditions can't be edited or added for now"

			Dim row1 As New TableRow(documentCore)
			If True Then
				Dim cell1 As New TableCell(documentCore)
				cell1.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Thick, SautinSoft.Document.Color.White, 1)
				cell1.CellFormat.BackgroundColor = New SautinSoft.Document.Color("#E6E6E6")
				cell1.CellFormat.Padding = New Padding(0, 3, 0, 0)

				Dim p1 As New Paragraph(documentCore, "Content Name")
				p1.ParagraphFormat.Style = TableContentStyle_1
				cell1.Blocks.Add(p1)
				row1.Cells.Add(cell1)
			End If
			If True Then
				Dim cell2 As New TableCell(documentCore)
				cell2.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Thick, SautinSoft.Document.Color.White, 1)
				cell2.CellFormat.BackgroundColor = New SautinSoft.Document.Color("#E6E6E6")
				cell2.CellFormat.Padding = New Padding(0, 3, 0, 0)
				Dim y = x.ToString().Replace(vbCrLf, "")
				TryCast(cell2, TableCell).Blocks.Content.Replace(y, SautinSoft.Document.LoadOptions.HtmlDefault)
				row1.Cells.Add(cell2)
			End If
			row1.RowFormat.Height = New TableRowHeight(15, HeightRule.Auto)
			table.Rows.Add(row1)
			Dim row2 As New TableRow(documentCore)
			If True Then
				Dim cell1 As New TableCell(documentCore)
				cell1.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Thick, SautinSoft.Document.Color.White, 1)
				cell1.CellFormat.BackgroundColor = New SautinSoft.Document.Color("#E6E6E6")
				cell1.CellFormat.Padding = New Padding(0, 3, 0, 0)
				Dim p2 As New Paragraph(documentCore, "Image")
				p2.ParagraphFormat.Style = TableContentStyle_1
				cell1.Blocks.Add(p2)
				row2.Cells.Add(cell1)
			End If
			If True Then
				Dim cell2 As New TableCell(documentCore)
				cell2.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Thick, SautinSoft.Document.Color.White, 1)
				cell2.CellFormat.BackgroundColor = New SautinSoft.Document.Color("#E6E6E6")
				cell2.CellFormat.Padding = New Padding(0, 3, 0, 0)
				Dim p3 As New Paragraph(documentCore, New Picture(documentCore, New MemoryStream(imageData)))
				p3.ParagraphFormat.Style = TableContentStyle_1
				cell2.Blocks.Add(p3)
				row2.Cells.Add(cell2)
			End If
			row2.RowFormat.Height = New TableRowHeight(15, HeightRule.Auto)
			table.Rows.Add(row2)
			Dim section As New Section(documentCore)
			section.PageSetup.PageColor.SetSolid(New SautinSoft.Document.Color("#f8f8fa"))
			section.PageSetup.PaperType = PaperType.A4
			section.PageSetup.Orientation = Orientation.Portrait
			section.PageSetup.PageMargins = New PageMargins() With {
				.Top = LengthUnitConverter.Convert(30, LengthUnit.Millimeter, LengthUnit.Point),
				.Right = LengthUnitConverter.Convert(20, LengthUnit.Millimeter, LengthUnit.Point),
				.Bottom = LengthUnitConverter.Convert(5, LengthUnit.Millimeter, LengthUnit.Point),
				.Left = LengthUnitConverter.Convert(0, LengthUnit.Millimeter, LengthUnit.Point)
			}
			section.Blocks.Add(table)

			documentCore.Sections.Add(section)
			Dim filePath As String = "structured-table.docx"
			documentCore.Save(filePath, New DocxSaveOptions())
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
