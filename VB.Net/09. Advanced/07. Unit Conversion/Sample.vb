Imports System
Imports System.IO
Imports SautinSoft.Document

Module Sample
    Sub Main()
        ConvertPointsTo()
        UsingUnitConversion()
    End Sub

    ''' <summary>
    ''' Shows what is the one point is equal in different units.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/unit-conversion.php
    ''' </remarks>
    Sub ConvertPointsTo()
        For Each unit As LengthUnit In System.Enum.GetValues(GetType(LengthUnit))
            Dim s As String = String.Format("1 point = {0} {1}", LengthUnitConverter.Convert(1, LengthUnit.Point, unit), unit.ToString().ToLowerInvariant())

            Console.WriteLine(s)
        Next unit
        Console.WriteLine("Press any key...")
        Console.ReadKey()
    End Sub

    ''' <summary>
    ''' How to convert between different units.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/unit-conversion.php
    ''' </remarks>
    Sub UsingUnitConversion()
        Dim dc As New DocumentCore()

        ' Add new section.
        Dim s1 As New Section(dc)
        s1.PageSetup.PageWidth = LengthUnitConverter.Convert(100, LengthUnit.Millimeter, LengthUnit.Point)
        s1.PageSetup.PageHeight = LengthUnitConverter.Convert(1, LengthUnit.Inch, LengthUnit.Point)
        s1.PageSetup.PageMargins.Left = LengthUnitConverter.Convert(0.3, LengthUnit.Centimeter, LengthUnit.Point)
        s1.PageSetup.PageMargins.Top = LengthUnitConverter.Convert(7, LengthUnit.Pixel, LengthUnit.Point)
        s1.PageSetup.PageMargins.Right = LengthUnitConverter.Convert(100, LengthUnit.Twip, LengthUnit.Point)
        s1.PageSetup.PageMargins.Bottom = LengthUnitConverter.Convert(5, LengthUnit.Millimeter, LengthUnit.Point)

        dc.Sections.Add(s1)
        s1.Content.Start.Insert("You are welcome!", New CharacterFormat() With {
                .FontColor = Color.Orange,
                .Size = 36
            })

        dc.Save("Result.docx")

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo("Result.docx") With {.UseShellExecute = True})
    End Sub
End Module