Imports Microsoft.VisualBasic
Imports System
Imports SautinSoft.Document
Imports SautinSoft.Document.Tables


Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			DetectBorders()
		End Sub
		''' <summary>
		''' Detect cell borders with the same color.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-detect-borders-in-table-csharp-vb-net.php
		''' </remarks>

		Private Shared Sub DetectBorders()
			Dim dc As DocumentCore = DocumentCore.Load("..\..\..\example.docx")

			For Each itemTC As TableCell In dc.GetChildElements(True, ElementType.TableCell)
				Dim sbLeft As SingleBorder = itemTC.CellFormat.Borders(SingleBorderType.Left)
				Dim sbTop As SingleBorder = itemTC.CellFormat.Borders(SingleBorderType.Top)
				Dim sbRight As SingleBorder = itemTC.CellFormat.Borders(SingleBorderType.Right)
				Dim sbBottom As SingleBorder = itemTC.CellFormat.Borders(SingleBorderType.Bottom)
				If sbLeft.Color = sbTop.Color AndAlso sbTop.Color = sbRight.Color AndAlso sbRight.Color = sbBottom.Color Then
					itemTC.Content.Start.Insert("This cell has the same border color." & vbCrLf)
				End If
			Next itemTC

			' Save our document into DOCX format.
			Dim filePath As String = "ResultDetectBorder.docx"
			dc.Save(filePath)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace
