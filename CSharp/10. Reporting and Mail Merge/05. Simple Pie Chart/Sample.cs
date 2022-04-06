using System;
using System.IO;
using System.Collections.Generic;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using System.Linq;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            PieChart();
        }
        /// <summary>
        /// Creates a simple Pie Chart in a document.
        /// </summary>
        /// <remarks>
        /// See details at: https://sautinsoft.com/products/document/help/net/developer-guide/reporting-create-simple-pie-chart-in-pdf-net-csharp-vb.php
        /// </remarks>
        public static void PieChart()
        {
            DocumentCore dc = new DocumentCore();

            FloatingLayout fl = new FloatingLayout(
                new HorizontalPosition(20f, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin),
                new VerticalPosition(15f, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin),
                new Size(200, 200));

            Dictionary<string, double> chartData = new Dictionary<string, double>
            {   {"Potato", 25},
                {"Carrot", 16},
                {"Salad", 10},
                {"Cucumber", 10},
                {"Tomato", 4},
                {"Rice", 30},
                {"Onion", 5} };

            AddPieChart(dc, fl, chartData, true, "%", true);


            // Let's save the document into PDF format (you may choose another).
            string filePath = @"Pie Chart.pdf";
            dc.Save(filePath);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
        }
        /// <summary>
        /// This method creates a pie chart.
        /// </summary>
        /// <param name="dc">Document</param>
        /// <param name="chartLayout">Chart layout</param>
        /// <param name="data">Chart data</param>
        /// <param name="addLabels">Add labels or not</param>
        /// <param name="labelSign">Label sign</param>
        /// <param name="addLegend">Add legend</param>
        /// <remarks>
        /// This method is made specially with open source code. You may change anything you want:<br />
        /// Chart colors, size, position, font name and font size, position of labels (inside or outside the chart), legend position.<br />
        /// If you need any assistance, feel free to email us: support@sautinsoft.com.<br />
        /// </remarks>
        public static void AddPieChart(DocumentCore dc, FloatingLayout chartLayout, Dictionary<string, double> data, bool addLabels = true, string labelSign = null, bool addLegend = true)
        {
            // Assume that our chart can have 10 pies maximally.
            // And we'll use these colors order by order.
            // You may change the colors and their quantity.
            List<string> colors = new List<string>()
            {
                "#70AD47", // light green
                "#4472C4", // blue
                "#FFC000", // yellow
                "#A5A5A5", // grey
                "#ED7D31", // orange
                "#5B9BD5", // light blue
                "#44546A", // blue and grey
                "#C00000", // red
                "#00B050", // green
                "#9933FF"  // purple

            };
            // 1. To create a circle chart, assume that the sum of all values are 100%
            // and calculate the percent for the each value.
            // Translate all data to perce
            double amount = data.Values.Sum();
            List<double> percentages = new List<double>();

            foreach (double v in data.Values)
                percentages.Add(v * 100 / amount);

            // 2. Translate the percentage value of the each pie into degrees.
            // The whole circle is 360 degrees.
            int pies = data.Values.Count;
            List<double> pieDegrees = new List<double>();
            foreach (double p in percentages)
                pieDegrees.Add(p * 360 / 100);

            // 3. Translate degrees to the "Pie" measurement.
            List<double> pieMeasure = new List<double>();
            // Add the start position.
            pieMeasure.Add(0);
            double currentAngle = 0;
            foreach (double pd in pieDegrees)
            {
                currentAngle += pd;
                pieMeasure.Add(480 * currentAngle / 360);
            }

            // 4. Create the pies.
            Shape originalShape = new Shape(dc, chartLayout);

            for (int i = 0; i < pies; i++)
            {
                Shape shpPie = originalShape.Clone(true);
                shpPie.Outline.Fill.SetSolid(Color.White);
                shpPie.Outline.Width = 0.5;
                shpPie.Fill.SetSolid(new Color(colors[i]));

                shpPie.Geometry.SetPreset(Figure.Pie);
                shpPie.Geometry.AdjustValues["adj1"] = 45000 * pieMeasure[i];
                shpPie.Geometry.AdjustValues["adj2"] = 45000 * pieMeasure[i + 1];

                dc.Content.End.Insert(shpPie.Content);
            }

            // 5. Add labels
            if (addLabels)
            {
                // 0.5 ... 1.2 (inside/outside the circle).
                double multiplier = 0.8;
                double radius = chartLayout.Size.Width / 2 * multiplier;
                currentAngle = 0;
                double labelW = 35;
                double labelH = 20;

                for (int i = 0; i < pieDegrees.Count; i++)
                {
                    currentAngle += pieDegrees[i];
                    double middleAngleDegrees = 360 - (currentAngle - pieDegrees[i] / 2);
                    double middleAngleRad = middleAngleDegrees * (Math.PI / 180);

                    // Calculate the (x, y) on the circle.
                    double x = radius * Math.Cos(middleAngleRad);
                    double y = radius * Math.Sin(middleAngleRad);

                    // Correct the position depending of the label size;
                    x -= labelW / 2;
                    y += labelH / 2;

                    HorizontalPosition centerH = new HorizontalPosition(chartLayout.HorizontalPosition.Value + chartLayout.Size.Width / 2, LengthUnit.Point, HorizontalPositionAnchor.LeftMargin);
                    VerticalPosition centerV = new VerticalPosition(chartLayout.VerticalPosition.Value + chartLayout.Size.Height / 2, LengthUnit.Point, VerticalPositionAnchor.TopMargin);

                    HorizontalPosition labelX = new HorizontalPosition(centerH.Value + x, LengthUnit.Point, HorizontalPositionAnchor.LeftMargin);
                    VerticalPosition labelY = new VerticalPosition(centerV.Value - y, LengthUnit.Point, VerticalPositionAnchor.TopMargin);

                    FloatingLayout labelLayout = new FloatingLayout(labelX, labelY, new Size(labelW, labelH));
                    Shape shpLabel = new Shape(dc, labelLayout);
                    shpLabel.Outline.Fill.SetEmpty();
                    shpLabel.Text.Blocks.Content.Start.Insert($"{data.Values.ElementAt(i)}{labelSign}",
                        new CharacterFormat()
                        {
                            FontName = "Arial",
                            Size = 10,
                            //FontColor = new Color("#333333")
                            FontColor = new Color("#FFFFFF")
                        });
                    (shpLabel.Text.Blocks[0] as Paragraph).ParagraphFormat.Alignment = HorizontalAlignment.Center;
                    dc.Content.End.Insert(shpLabel.Content);
                }

                // 6. Add Legend
                if (addLegend)
                {
                    double legendTopMargin = LengthUnitConverter.Convert(5, LengthUnit.Millimeter, LengthUnit.Point);
                    double legendLeftMargin = LengthUnitConverter.Convert(-10, LengthUnit.Millimeter, LengthUnit.Point);
                    HorizontalPosition legendX = new HorizontalPosition(chartLayout.HorizontalPosition.Value + legendLeftMargin, LengthUnit.Point, HorizontalPositionAnchor.LeftMargin);
                    VerticalPosition legendY = new VerticalPosition(chartLayout.VerticalPosition.Value + chartLayout.Size.Height + legendTopMargin, LengthUnit.Point, VerticalPositionAnchor.TopMargin);
                    double legendW = chartLayout.Size.Width * 2;
                    double legendH = 20;

                    FloatingLayout legendLayout = new FloatingLayout(legendX, legendY, new Size(legendW, legendH));
                    Shape shpLegend = new Shape(dc, legendLayout);
                    shpLegend.Outline.Fill.SetEmpty();

                    Paragraph pLegend = new Paragraph(dc);
                    pLegend.ParagraphFormat.Alignment = HorizontalAlignment.Left;

                    for (int i = 0; i < data.Count; i++)
                    {
                        string legendItem = data.Keys.ElementAt(i);

                        // 183 - circle, "Symbol"
                        // 167 - square, "Wingdings"

                        Run marker = new Run(dc, (char)167, new CharacterFormat()
                        {
                            FontColor = new Color(colors[i]),
                            FontName = "Wingdings"
                        });
                        pLegend.Content.End.Insert(marker.Content);
                        pLegend.Content.End.Insert($" {legendItem}   ", new CharacterFormat());
                    }


                    shpLegend.Text.Blocks.Add(pLegend);
                    dc.Content.End.Insert(shpLegend.Content);
                }
            }
        }

    }
}
