Imports System
Imports System.Data
Imports System.IO
Imports SautinSoft.Document

Namespace Sample
    Friend Class Sample
        Shared Sub Main(ByVal args() As String)
            TableReportWithRegions()
        End Sub

        ''' <summary>
        ''' Generates a table report with regions using XML document as a data source.
        ''' </summary>
        ''' <remarks>
        ''' See details at: https://www.sautinsoft.com/products/document/help/net/developer-guide/mail-merge-table-report-with-regions-net-csharp-vb.php
        ''' </remarks>
        Public Shared Sub TableReportWithRegions()
            ' Create the Dataset and read the XML.
            Dim ds As New DataSet()

            ds.ReadXml("..\..\..\Orders.xml")

            ' Load the template document.
            Dim templatePath As String = "..\..\..\InvoiceTemplate.docx"

            Dim dc As DocumentCore = DocumentCore.Load(templatePath)

            ' Execute the mail merge.
            dc.MailMerge.Execute(ds.Tables("Order"))

            Dim resultPath As String = "Invoices.pdf"

            ' Save the output to file
            dc.Save(resultPath)

            ' Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(resultPath) With {.UseShellExecute = True})
        End Sub
    End Class
End Namespace
