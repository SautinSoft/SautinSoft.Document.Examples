Imports System
Imports SautinSoft.Document
Imports SautinSoft.Document.Tables
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports SautinSoft.Document.Drawing

Namespace Example
	Friend Class Program
		Shared Sub Main(ByVal args() As String)
			TableVariousBorders()
		End Sub
		''' <summary>
		''' The simple report using DocumentBuilder and saves it in a desired format.
		''' </summary>
		''' <remarks>
		''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-create-document-with-various-table-borders-in-csharp-vb-net.php
		''' </remarks>

		Private Shared Sub TableVariousBorders()
			Dim dc As New DocumentCore()
			Dim db As New DocumentBuilder(dc)

			' The Picture for visualization.
			Dim logo As Picture = db.InsertImage("..\..\..\medical.png", New Size(50, 25, LengthUnit.Millimeter))
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)

			' Insert the formatted text into the document using DocumentBuilder.
			db.CharacterFormat.FontName = "Verdana"
			db.CharacterFormat.Size = 16.5F
			db.CharacterFormat.AllCaps = True
			db.CharacterFormat.Italic = True
			db.CharacterFormat.FontColor = Color.Orange

			db.ParagraphFormat.Alignment = HorizontalAlignment.Center
			db.Writeln("Outpatient card")

			' This method will clear all directly set formatting values.
			db.ParagraphFormat.ClearFormatting()
			db.CharacterFormat.ClearFormatting()

			' Data Source: SQL, DB, ACCESS, Data Array, etc.
			Dim namepatient As String = "Smith John"
			Dim namedoctor As String = "Dr. Christopher Duncan Turk"
			Dim dateofreceipt As String = "2021/06/03"
			Dim dischargedate As String = "2021/06/24"
			Dim year As String = "1961/03/03"
			Dim address As String = "1775 Westminster Avenue APT D56 Brooklyn, NY, 11235"
			Dim diagnosis As String = "Pneumonia (Covid19)."
			Dim treatment As String = "Dexamethasone Remdesivir Anticoagulation drugs."

			' Create a new table with preferred width.
			Dim maintable As Table = db.StartTable()
			db.TableFormat.PreferredWidth = New TableWidth(LengthUnitConverter.Convert(7, LengthUnit.Inch, LengthUnit.Point), TableWidthUnit.Point)

			' Specify formatting of cells and alignment.
			db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.DotDotDash, Color.Green, 1)
			db.CellFormat.BackgroundColor = Color.DarkYellow
			db.CellFormat.VerticalAlignment = VerticalAlignment.Center
			db.ParagraphFormat.Alignment = HorizontalAlignment.Center

			' Specify height of rows and write text.
			db.RowFormat.Height = New TableRowHeight(45.0F, HeightRule.Auto)

			db.InsertCell()
			db.Write("First/Last Name")
			db.InsertCell()
			db.Write("Date of Birth ")
			db.InsertCell()
			db.Write("Address")
			db.InsertCell()
			db.Write("Diagnosis")
			db.InsertCell()
			db.Write("Treatment")
			db.EndRow()

			' Specify formatting of cells and alignment.
			db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.DoubleWave, Color.Black, 1)
			db.CellFormat.BackgroundColor = Color.LightGray
			db.CellFormat.VerticalAlignment = VerticalAlignment.Center
			db.ParagraphFormat.Alignment = HorizontalAlignment.Left

			' Specify height of rows and write text.
			db.RowFormat.Height = New TableRowHeight(100.0F, HeightRule.Exact)
			db.InsertCell()
			db.Write(namepatient)
			db.InsertCell()
			db.Write(year)
			db.InsertCell()
			db.Write(address)
			db.InsertCell()
			db.Write(diagnosis)
			db.InsertCell()
			db.Write(treatment)
			db.EndRow()
			db.EndTable()

			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)

			' Insert the formatted text into the document using DocumentBuilder.
			db.CharacterFormat.FontName = "Calibri"
			db.CharacterFormat.Size = 13.5F
			db.CharacterFormat.FontColor = Color.DarkGreen
			db.ParagraphFormat.SpecialIndentation = 100
			db.Writeln("*The card is filled in by the attending physician ")
			db.ParagraphFormat.ClearFormatting()
			db.CharacterFormat.ClearFormatting()
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)

			Dim addtable As Table = db.StartTable()
			db.TableFormat.IndentFromLeft = 200
			db.TableFormat.PreferredWidth = New TableWidth(LengthUnitConverter.Convert(4, LengthUnit.Inch, LengthUnit.Point), TableWidthUnit.Point)

			' Specify formatting of cells and alignment.
			db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dashed, Color.Orange, 2)
			db.CellFormat.BackgroundColor = Color.Magenta
			db.CellFormat.VerticalAlignment = VerticalAlignment.Center
			db.ParagraphFormat.Alignment = HorizontalAlignment.Center

			' Specify height of rows and write text.
			db.RowFormat.Height = New TableRowHeight(15.0F, HeightRule.Exact)

			db.InsertCell()
			db.Write("Attending doctor")
			db.InsertCell()
			db.Write("Date of receipt")
			db.InsertCell()
			db.Write("Discharge date")
			db.EndRow()

			' Specify formatting of cells and alignment.
			db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dotted, Color.Purple, 2)
			db.CellFormat.BackgroundColor = Color.LightGray
			db.CellFormat.VerticalAlignment = VerticalAlignment.Top
			db.ParagraphFormat.Alignment = HorizontalAlignment.Center

			' Specify height of rows and write text.
			db.RowFormat.Height = New TableRowHeight(25.0F, HeightRule.Exact)
			db.InsertCell()
			db.Write(namedoctor)
			db.InsertCell()
			db.Write(dateofreceipt)
			db.InsertCell()
			db.Write(dischargedate)
			db.EndRow()
			db.EndTable()

			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)
			db.InsertSpecialCharacter(SpecialCharacterType.LineBreak)

			db.ParagraphFormat.Alignment = HorizontalAlignment.Right
			Dim sign As Picture = db.InsertImage("..\..\..\sign.png", New Size(50, 25, LengthUnit.Millimeter))

			' Save our document into PDF format.
			Dim filePath As String = "Result.pdf"
			dc.Save(filePath)

			' Open the result for demonstration purposes.
			System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
		End Sub
	End Class
End Namespace