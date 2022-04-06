using System;
using SautinSoft.Document;
using SautinSoft.Document.Tables;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SautinSoft.Document.Drawing;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            TableVariousBorders();
        }
        /// <summary>
        /// The simple report using DocumentBuilder and saves it in a desired format.
        /// </summary>
        /// <remarks>
        /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/from-customers-create-document-with-various-table-borders-in-csharp-vb-net.php
        /// </remarks>

        static void TableVariousBorders()
        {
            DocumentCore dc = new DocumentCore();
            DocumentBuilder db = new DocumentBuilder(dc);

            // The Picture for visualization.
            Picture logo = db.InsertImage(@"..\..\..\medical.png", new Size(50, 25, LengthUnit.Millimeter));
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);

            // Insert the formatted text into the document using DocumentBuilder.
            db.CharacterFormat.FontName = "Verdana";
            db.CharacterFormat.Size = 16.5f;
            db.CharacterFormat.AllCaps = true;
            db.CharacterFormat.Italic = true;
            db.CharacterFormat.FontColor = Color.Orange;

            db.ParagraphFormat.Alignment = HorizontalAlignment.Center;
            db.Writeln("Outpatient card");

            // This method will clear all directly set formatting values.
            db.ParagraphFormat.ClearFormatting();
            db.CharacterFormat.ClearFormatting();

            // Data Source: SQL, DB, ACCESS, Data Array, etc.
            string namepatient = "Smith John";
            string namedoctor = "Dr. Christopher Duncan Turk";
            string dateofreceipt = "2021/06/03";
            string dischargedate = "2021/06/24";
            string year = "1961/03/03";
            string address = "1775 Westminster Avenue APT D56 Brooklyn, NY, 11235";
            string diagnosis = "Pneumonia (Covid19).";
            string treatment = "Dexamethasone Remdesivir Anticoagulation drugs.";

            // Create a new table with preferred width.
            Table maintable = db.StartTable();
            db.TableFormat.PreferredWidth = new TableWidth(LengthUnitConverter.Convert(7, LengthUnit.Inch, LengthUnit.Point), TableWidthUnit.Point);

            // Specify formatting of cells and alignment.
            db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.DotDotDash, Color.Green, 1);
            db.CellFormat.BackgroundColor = Color.DarkYellow;
            db.CellFormat.VerticalAlignment = VerticalAlignment.Center;
            db.ParagraphFormat.Alignment = HorizontalAlignment.Center;

            // Specify height of rows and write text.
            db.RowFormat.Height = new TableRowHeight(45f, HeightRule.Auto);

            db.InsertCell();
            db.Write("First/Last Name");
            db.InsertCell();
            db.Write("Date of Birth ");
            db.InsertCell();
            db.Write("Address");
            db.InsertCell();
            db.Write("Diagnosis");
            db.InsertCell();
            db.Write("Treatment");
            db.EndRow();

            // Specify formatting of cells and alignment.
            db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.DoubleWave, Color.Black, 1);
            db.CellFormat.BackgroundColor = Color.LightGray;
            db.CellFormat.VerticalAlignment = VerticalAlignment.Center;
            db.ParagraphFormat.Alignment = HorizontalAlignment.Left;

            // Specify height of rows and write text.
            db.RowFormat.Height = new TableRowHeight(100f, HeightRule.Exact);
            db.InsertCell();
            db.Write(namepatient);
            db.InsertCell();
            db.Write(year);
            db.InsertCell();
            db.Write(address);
            db.InsertCell();
            db.Write(diagnosis);
            db.InsertCell();
            db.Write(treatment);
            db.EndRow();
            db.EndTable();

            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);

            // Insert the formatted text into the document using DocumentBuilder.
            db.CharacterFormat.FontName = "Calibri";
            db.CharacterFormat.Size = 13.5f;
            db.CharacterFormat.FontColor = Color.DarkGreen;
            db.ParagraphFormat.SpecialIndentation = 100;
            db.Writeln("*The card is filled in by the attending physician ");
            db.ParagraphFormat.ClearFormatting();
            db.CharacterFormat.ClearFormatting();
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);

            Table addtable = db.StartTable();
            db.TableFormat.IndentFromLeft = 200;
            db.TableFormat.PreferredWidth = new TableWidth(LengthUnitConverter.Convert(4, LengthUnit.Inch, LengthUnit.Point), TableWidthUnit.Point);

            // Specify formatting of cells and alignment.
            db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dashed, Color.Orange, 2);
            db.CellFormat.BackgroundColor = Color.Magenta;
            db.CellFormat.VerticalAlignment = VerticalAlignment.Center;
            db.ParagraphFormat.Alignment = HorizontalAlignment.Center;

            // Specify height of rows and write text.
            db.RowFormat.Height = new TableRowHeight(15f, HeightRule.Exact);

            db.InsertCell();
            db.Write("Attending doctor");
            db.InsertCell();
            db.Write("Date of receipt");
            db.InsertCell();
            db.Write("Discharge date");
            db.EndRow();

            // Specify formatting of cells and alignment.
            db.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Dotted, Color.Purple, 2);
            db.CellFormat.BackgroundColor = Color.LightGray;
            db.CellFormat.VerticalAlignment = VerticalAlignment.Top;
            db.ParagraphFormat.Alignment = HorizontalAlignment.Center;

            // Specify height of rows and write text.
            db.RowFormat.Height = new TableRowHeight(25f, HeightRule.Exact);
            db.InsertCell();
            db.Write(namedoctor);
            db.InsertCell();
            db.Write(dateofreceipt);
            db.InsertCell();
            db.Write(dischargedate);
            db.EndRow();
            db.EndTable();

            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);
            db.InsertSpecialCharacter(SpecialCharacterType.LineBreak);

            db.ParagraphFormat.Alignment = HorizontalAlignment.Right;
            Picture sign = db.InsertImage(@"..\..\..\sign.png", new Size(50, 25, LengthUnit.Millimeter));

            // Save our document into PDF format.
            string filePath = "Result.pdf";
            dc.Save(filePath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}