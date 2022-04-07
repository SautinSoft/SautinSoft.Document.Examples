Imports System
Imports System.IO
Imports System.Linq
Imports System.Collections.Generic
Imports SautinSoft.Document
Imports SautinSoft.Document.Drawing

Module Sample
    Sub Main()
        PieChart()
    End Sub

    ''' <summary>
    ''' Creates a simple Pie Chart in a document.
    ''' </summary>
    ''' <remarks>
    ''' See details at: https://sautinsoft.com/products/document/help/net/developer-guide/reporting-create-simple-pie-chart-in-pdf-net-csharp-vb.php
    ''' </remarks>
    Public Sub PieChart()
        Dim dc As New DocumentCore()

        Dim fl As New FloatingLayout(New HorizontalPosition(20.0F, LengthUnit.Millimeter, HorizontalPositionAnchor.LeftMargin), New VerticalPosition(15.0F, LengthUnit.Millimeter, VerticalPositionAnchor.TopMargin), New Size(200, 200))

        Dim chartData As New Dictionary(Of String, Double) From {
                {"Potato", 25},
                {"Carrot", 16},
                {"Salad", 10},
                {"Cucumber", 10},
                {"Tomato", 4},
                {"Rice", 30},
                {"Onion", 5}
            }

        AddPieChart(dc, fl, chartData, True, "%", True)


        ' Let's save the document into PDF format (you may choose another).
        Dim filePath As String = "Pie Chart.pdf"
        dc.Save(filePath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub
    ''' <summary>
    ''' This method creates a pie chart.
    ''' </summary>
    ''' <param name="dc">Document</param>
    ''' <param name="chartLayout">Chart layout</param>
    ''' <param name="data">Chart data</param>
    ''' <param name="addLabels">Add labels or not</param>
    ''' <param name="labelSign">Label sign</param>
    ''' <param name="addLegend">Add legend</param>
    ''' <remarks>
    ''' This method is made specially with open source code. You may change anything you want:<br />
    ''' Chart colors, size, position, font name and font size, position of labels (inside or outside the chart), legend position.<br />
    ''' If you need any assistance, feel free to email us: support@sautinsoft.com.<br />
    ''' </remarks>
    Public Sub AddPieChart(ByVal dc As DocumentCore, ByVal chartLayout As FloatingLayout, ByVal data As Dictionary(Of String, Double), Optional ByVal addLabels As Boolean = True, Optional ByVal labelSign As String = Nothing, Optional ByVal addLegend As Boolean = True)
        ' Assume that our chart can have 10 pies maximally.
        ' And we'll use these colors order by order.
        ' You may change the colors and their quantity.
        Dim colors As New List(Of String)() From {"#70AD47", "#4472C4", "#FFC000", "#A5A5A5", "#ED7D31", "#5B9BD5", "#44546A", "#C00000", "#00B050", "#9933FF"}
        ' 1. To create a circle chart, assume that the sum of all values are 100%
        ' and calculate the percent for the each value.
        ' Translate all data to perce
        Dim amount As Double = data.Values.Sum()
        Dim percentages As New List(Of Double)()

        For Each v As Double In data.Values
            percentages.Add(v * 100 / amount)
        Next v

        ' 2. Translate the percentage value of the each pie into degrees.
        ' The whole circle is 360 degrees.
        Dim pies As Integer = data.Values.Count
        Dim pieDegrees As New List(Of Double)()
        For Each p As Double In percentages
            pieDegrees.Add(p * 360 / 100)
        Next p

        ' 3. Translate degrees to the "Pie" measurement.
        Dim pieMeasure As New List(Of Double)()
        ' Add the start position.
        pieMeasure.Add(0)
        Dim currentAngle As Double = 0
        For Each pd As Double In pieDegrees
            currentAngle += pd
            pieMeasure.Add(480 * currentAngle / 360)
        Next pd

        ' 4. Create the pies.
        Dim originalShape As New Shape(dc, chartLayout)

        For i As Integer = 0 To pies - 1
            Dim shpPie As Shape = originalShape.Clone(True)
            shpPie.Outline.Fill.SetSolid(Color.White)
            shpPie.Outline.Width = 0.5
            shpPie.Fill.SetSolid(New Color(colors(i)))

            shpPie.Geometry.SetPreset(Figure.Pie)
            shpPie.Geometry.AdjustValues("adj1") = 45000 * pieMeasure(i)
            shpPie.Geometry.AdjustValues("adj2") = 45000 * pieMeasure(i + 1)

            dc.Content.End.Insert(shpPie.Content)
        Next i

        ' 5. Add labels
        If addLabels Then
            ' 0.5 ... 1.2 (inside/outside the circle).
            Dim multiplier As Double = 0.8
            Dim radius As Double = chartLayout.Size.Width / 2 * multiplier
            currentAngle = 0
            Dim labelW As Double = 35
            Dim labelH As Double = 20

            For i As Integer = 0 To pieDegrees.Count - 1
                currentAngle += pieDegrees(i)
                Dim middleAngleDegrees As Double = 360 - (currentAngle - pieDegrees(i) / 2)
                Dim middleAngleRad As Double = middleAngleDegrees * (Math.PI / 180)

                ' Calculate the (x, y) on the circle.
                Dim x As Double = radius * Math.Cos(middleAngleRad)
                Dim y As Double = radius * Math.Sin(middleAngleRad)

                ' Correct the position depending of the label size;
                x -= labelW / 2
                y += labelH / 2

                Dim centerH As New HorizontalPosition(chartLayout.HorizontalPosition.Value + chartLayout.Size.Width \ 2, LengthUnit.Point, HorizontalPositionAnchor.LeftMargin)
                Dim centerV As New VerticalPosition(chartLayout.VerticalPosition.Value + chartLayout.Size.Height \ 2, LengthUnit.Point, VerticalPositionAnchor.TopMargin)

                Dim labelX As New HorizontalPosition(centerH.Value + x, LengthUnit.Point, HorizontalPositionAnchor.LeftMargin)
                Dim labelY As New VerticalPosition(centerV.Value - y, LengthUnit.Point, VerticalPositionAnchor.TopMargin)

                Dim labelLayout As New FloatingLayout(labelX, labelY, New Size(labelW, labelH))
                Dim shpLabel As New Shape(dc, labelLayout)
                shpLabel.Outline.Fill.SetEmpty()
                shpLabel.Text.Blocks.Content.Start.Insert($"{data.Values.ElementAt(i)}{labelSign}", New CharacterFormat() With {
                        .FontName = "Arial",
                        .Size = 10,
                        .FontColor = New Color("#FFFFFF")
                    })
                TryCast(shpLabel.Text.Blocks(0), Paragraph).ParagraphFormat.Alignment = HorizontalAlignment.Center
                dc.Content.End.Insert(shpLabel.Content)
            Next i

            ' 6. Add Legend
            If addLegend Then
                Dim legendTopMargin As Double = LengthUnitConverter.Convert(5, LengthUnit.Millimeter, LengthUnit.Point)
                Dim legendLeftMargin As Double = LengthUnitConverter.Convert(-10, LengthUnit.Millimeter, LengthUnit.Point)
                Dim legendX As New HorizontalPosition(chartLayout.HorizontalPosition.Value + legendLeftMargin, LengthUnit.Point, HorizontalPositionAnchor.LeftMargin)
                Dim legendY As New VerticalPosition(chartLayout.VerticalPosition.Value + chartLayout.Size.Height + legendTopMargin, LengthUnit.Point, VerticalPositionAnchor.TopMargin)
                Dim legendW As Double = chartLayout.Size.Width * 2
                Dim legendH As Double = 20

                Dim legendLayout As New FloatingLayout(legendX, legendY, New Size(legendW, legendH))
                Dim shpLegend As New Shape(dc, legendLayout)
                shpLegend.Outline.Fill.SetEmpty()

                Dim pLegend As New Paragraph(dc)
                pLegend.ParagraphFormat.Alignment = HorizontalAlignment.Left

                For i As Integer = 0 To data.Count - 1
                    Dim legendItem As String = data.Keys.ElementAt(i)

                    ' 183 - circle, "Symbol"
                    ' 167 - square, "Wingdings"

                    Dim marker As Run = New Run(dc, ChrW(167), New CharacterFormat() With {
                            .FontColor = New Color(colors(i)),
                            .FontName = "Wingdings"
                        })
                    pLegend.Content.End.Insert(marker.Content)
                    pLegend.Content.End.Insert($" {legendItem}   ", New CharacterFormat())
                Next i


                shpLegend.Text.Blocks.Add(pLegend)
                dc.Content.End.Insert(shpLegend.Content)
            End If
        End If
    End Sub

End Module