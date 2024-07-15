using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using SautinSoft.Document.Tables;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get your free 100-day key here:   
            // https://sautinsoft.com/start-for-free/

				StructuredTable();
		}
        /// <summary>
        /// Creating a structured table with data of different formats.
        /// </summary>
        /// <remarks>
        /// Details: https://sautinsoft.com/products/document/help/net/developer-guide/create-structured-table.php
        /// </remarks>
        static void StructuredTable()
        {		
            DocumentCore documentCore = new DocumentCore();
            byte[] imageData = File.ReadAllBytes(@"../../../image/smile.jpg");
            ParagraphStyle TableHeaderStyle = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Normal, documentCore);
            TableHeaderStyle.Name = "TableHeaderStyle";
            TableHeaderStyle.ParagraphFormat.Alignment = HorizontalAlignment.Left;
            TableHeaderStyle.ParagraphFormat.LeftIndentation = 6;
            TableHeaderStyle.CharacterFormat.FontColor = SautinSoft.Document.Color.White;
            TableHeaderStyle.CharacterFormat.FontName = "Segoe UI";
            TableHeaderStyle.CharacterFormat.Size = 9;
            documentCore.Styles.Add(TableHeaderStyle);

            ParagraphStyle TableContentStyle_1 = (ParagraphStyle)Style.CreateStyle(StyleTemplateType.Normal, documentCore);
            TableContentStyle_1.Name = "TableContentStyle_1";
            TableContentStyle_1.ParagraphFormat.Alignment = HorizontalAlignment.Left;
            TableContentStyle_1.ParagraphFormat.LeftIndentation = 6;
            //TableContentStyle_1.CharacterFormat.FontName = "Segoe UI";
            TableContentStyle_1.CharacterFormat.Size = 9;
            documentCore.Styles.Add(TableContentStyle_1);

            TableStyle TableOutlineStyle = (TableStyle)Style.CreateStyle(StyleTemplateType.TableNormal, documentCore);
            TableOutlineStyle.Name = "TableOutlineStyle";
            TableOutlineStyle.ParagraphFormat.SpaceAfter = 10;
            documentCore.Styles.Add(TableOutlineStyle);

            Table table = new Table(documentCore);
            table.TableFormat.PreferredWidth = new TableWidth(100, TableWidthUnit.Percentage);
            table.TableFormat.AutomaticallyResizeToFitContents = true;
            table.TableFormat.Alignment = HorizontalAlignment.Left;
            table.TableFormat.Style = TableOutlineStyle;

            TableRow headerRow = new TableRow(documentCore);
            {
                TableCell cell1 = new TableCell(documentCore);
                cell1.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Thick, SautinSoft.Document.Color.White, 1.5);
                cell1.CellFormat.BackgroundColor = new SautinSoft.Document.Color("#0054A6");
                cell1.CellFormat.Padding = new Padding(0, 3, 0, 0);
                Paragraph pHeader = new Paragraph(documentCore, "HeaderName");
                pHeader.ParagraphFormat.Style = TableHeaderStyle;
                cell1.Blocks.Add(pHeader);
                headerRow.Cells.Add(cell1);
            }
            {
                TableCell cell2 = new TableCell(documentCore);
                cell2.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Thick, SautinSoft.Document.Color.White, 1.5);
                cell2.CellFormat.BackgroundColor = new SautinSoft.Document.Color("#0054A6");
                cell2.CellFormat.Padding = new Padding(0, 3, 0, 0);
                Paragraph pHeaderContent = new Paragraph(documentCore, "HeaderContent");
                pHeaderContent.ParagraphFormat.Style = TableHeaderStyle;
                cell2.Blocks.Add(pHeaderContent);
                headerRow.Cells.Add(cell2);
            }
            headerRow.RowFormat.Height = new TableRowHeight(15, HeightRule.Auto);
            headerRow.RowFormat.RepeatOnEachPage = true;
            table.Rows.Add(headerRow);

            string x = "Detects if your application experiences an abnormal rise in the rate of HTTP requests or dependency calls that are reported as failed." +
                " The anomaly detection uses machine learning algorithms and occurs in near real time, " +
                "therefore there's no need to define a frequency for this signal.<br><br>To help you triage and diagnose the problem, an analysis of the characteristics of the failures and related telemetry is provided with the detection. " +
                "This feature works for any app, hosted in the cloud or on your own servers, that generates request or dependency telemetry - for example, " +
                "if you have a worker role that calls <a class=\"ext-smartDetecor-link\" href=\"https://docs.microsoft.com/azure/application-insights/app-insights-api-custom-events-metrics#trackrequest\" target=\"_blank\">TrackRequest()</a> " +
                "or <a class=\"ext-smartDetecor-link\" href=\"https://docs.microsoft.com/azure/application-insights/app-insights-api-custom-events-metrics#trackdependency\" target=\"_blank\">TrackDependency()</a>." +
                "<br/><br/><a class=\"ext-smartDetecor-link\" href=\"https://docs.microsoft.com/azure/azure-monitor/app/proactive-failure-diagnostics\" target=\"_blank\">Learn more about Failure Anomalies</a><br><br>" +
                "<p style=\"font-size: 13px; font-weight: 700;\">A note about your data privacy:</p><br><br>The service is entirely automatic and only you can see these notifications. " +
                "<a class=\"ext-smartDetecor-link\" href=\"https://docs.microsoft.com/en-us/azure/azure-monitor/app/data-retention-privacy\" target=\"_blank\">Read more about data privacy</a><br><br>Smart Alerts conditions can't be edited or added for now";

            TableRow row1 = new TableRow(documentCore);
            {
                TableCell cell1 = new TableCell(documentCore);
                cell1.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Thick, SautinSoft.Document.Color.White, 1);
                cell1.CellFormat.BackgroundColor = new SautinSoft.Document.Color("#E6E6E6");
                cell1.CellFormat.Padding = new Padding(0, 3, 0, 0);

                Paragraph p1 = new Paragraph(documentCore, "Content Name");
                p1.ParagraphFormat.Style = TableContentStyle_1;
                cell1.Blocks.Add(p1);
                row1.Cells.Add(cell1);
            }
            {
                TableCell cell2 = new TableCell(documentCore);
                cell2.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Thick, SautinSoft.Document.Color.White, 1);
                cell2.CellFormat.BackgroundColor = new SautinSoft.Document.Color("#E6E6E6");
                cell2.CellFormat.Padding = new Padding(0, 3, 0, 0);
                var y = x.ToString().Replace("\r\n", "");
                (cell2 as TableCell).Blocks.Content.Replace(y, SautinSoft.Document.LoadOptions.HtmlDefault); 
                row1.Cells.Add(cell2);
            }
            row1.RowFormat.Height = new TableRowHeight(15, HeightRule.Auto);
            table.Rows.Add(row1);
            TableRow row2 = new TableRow(documentCore);
            {
                TableCell cell1 = new TableCell(documentCore);
                cell1.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Thick, SautinSoft.Document.Color.White, 1);
                cell1.CellFormat.BackgroundColor = new SautinSoft.Document.Color("#E6E6E6");
                cell1.CellFormat.Padding = new Padding(0, 3, 0, 0);
                Paragraph p2 = new Paragraph(documentCore, "Image");
                p2.ParagraphFormat.Style = TableContentStyle_1;
                cell1.Blocks.Add(p2);
                row2.Cells.Add(cell1);
            }
            {
                TableCell cell2 = new TableCell(documentCore);
                cell2.CellFormat.Borders.SetBorders(MultipleBorderTypes.Outside, BorderStyle.Thick, SautinSoft.Document.Color.White, 1);
                cell2.CellFormat.BackgroundColor = new SautinSoft.Document.Color("#E6E6E6");
                cell2.CellFormat.Padding = new Padding(0, 3, 0, 0);
                Paragraph p3 = new Paragraph(documentCore, new Picture(documentCore, new MemoryStream(imageData)));
                p3.ParagraphFormat.Style = TableContentStyle_1;
                cell2.Blocks.Add(p3);
                row2.Cells.Add(cell2);
            }
            row2.RowFormat.Height = new TableRowHeight(15, HeightRule.Auto);
            table.Rows.Add(row2);
            Section section = new Section(documentCore);
            section.PageSetup.PageColor.SetSolid(new SautinSoft.Document.Color("#f8f8fa"));
            section.PageSetup.PaperType = PaperType.A4;
            section.PageSetup.Orientation = Orientation.Portrait;
            section.PageSetup.PageMargins = new PageMargins()
            {
                Top = LengthUnitConverter.Convert(30, LengthUnit.Millimeter, LengthUnit.Point),
                Right = LengthUnitConverter.Convert(20, LengthUnit.Millimeter, LengthUnit.Point),
                Bottom = LengthUnitConverter.Convert(5, LengthUnit.Millimeter, LengthUnit.Point),
                Left = LengthUnitConverter.Convert(0, LengthUnit.Millimeter, LengthUnit.Point)
            };
            section.Blocks.Add(table);

            documentCore.Sections.Add(section);
            string filePath = @"structured-table.docx";
            documentCore.Save(filePath, new DocxSaveOptions { });
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
    }
}