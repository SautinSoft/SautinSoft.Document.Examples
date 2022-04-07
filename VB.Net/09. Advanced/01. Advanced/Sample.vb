Imports SautinSoft.Document
Imports System.Linq


Module Sample
    Sub Main()
        FormDropDown()
    End Sub

    ''' <summary>
    ''' Creates a document containing FormDropDown element.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://sautinsoft.com/products/document/help/net/developer-guide/advanced.php
    ''' </remarks>
    Sub FormDropDown()
        Dim filePath As String = "Advanced.pdf"

        ' Let's create document.
        Dim dc As New DocumentCore()
        dc.Content.End.Insert((New Paragraph(dc, "The paragraph with FormDropDown element: ")).Content)
        Dim par As Paragraph = TryCast(dc.GetChildElements(True, ElementType.Paragraph).FirstOrDefault(), Paragraph)

        Dim field As FormDropDownData = TryCast((New Field(dc, FieldType.FormDropDown)).FormData, FormDropDownData)
        field.Items.Add("First Item")
        field.Items.Add("Second Item")
        field.Items.Add("Third Item")
        field.SelectedItemIndex = 2

        par.Inlines.Add(field.Field)

        ' Save our document.
        dc.Save(filePath)

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub
End Module