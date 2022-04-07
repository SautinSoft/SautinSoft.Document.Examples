Option Infer On

Imports System
Imports System.Globalization
Imports System.Text
Imports SautinSoft.Document
Imports SautinSoft.Document.MailMerging
Imports SautinSoft.Document.Tables

Module Sample
    Sub Main()
        FormsAndFields()
    End Sub

    ''' <summary>
    ''' Generate document (PDF) with forms and fields.
    ''' </summary>
    ''' <remarks>
    ''' Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/forms-and-fields.php
    ''' </remarks>
    Sub FormsAndFields()
        Dim dc As New DocumentCore()

        Dim placeHolder As New String(ChrW(&H2002), 50)

        ' Create form fields.
        Dim fFullName As New Field(dc, FieldType.FormText, Nothing, placeHolder)
        fFullName.FormData.Name = "FullName"
        fFullName.FormData.Enabled = True


        Dim fBirthData As New Field(dc, FieldType.FormText, Nothing, placeHolder)
        fBirthData.FormData.Name = "BirthDate"

        Dim fGender As New Field(dc, FieldType.FormDropDown)
        fGender.FormData.Name = "Gender"

        Dim fMarried As New Field(dc, FieldType.FormCheckBox)
        fMarried.FormData.Name = "Married"
        fMarried.FormData.Enabled = True

        Dim fPhone As New Field(dc, FieldType.FormText, Nothing, placeHolder)
        fPhone.FormData.Name = "Phone"

        dc.Sections.Add(New Section(dc, New Paragraph(dc, New Run(dc, "Full name: "), fFullName), New Paragraph(dc, New Run(dc, "Birth date: "), fBirthData), New Paragraph(dc, New Run(dc, "Gender: "), fGender), New Paragraph(dc, New Run(dc, "Married: "), fMarried), New Paragraph(dc, New Run(dc, "Phone: "), fPhone)))


        ' Customize form fields.
        Dim formFieldsData = dc.Content.FormFieldsData

        Dim fullNameFieldData = CType(formFieldsData("FullName"), FormTextData)
        fullNameFieldData.MaximumLength = 50
        fullNameFieldData.HelpText = "Enter your name and surname (trimmed to 50 characters)."
        fullNameFieldData.StatusText = fullNameFieldData.HelpText
        fullNameFieldData.Field.ResultInlines.Content.Replace("Mister Bean")

        Dim birthdateFieldData = CType(formFieldsData("BirthDate"), FormTextData)
        birthdateFieldData.TextType = FormTextType.Date
        birthdateFieldData.DefaultValue = "1990-01-01"
        birthdateFieldData.ValueFormat = "yyyy-MM-dd"
        birthdateFieldData.HelpText = "Enter your date of birth."
        birthdateFieldData.StatusText = birthdateFieldData.HelpText
        birthdateFieldData.Field.ResultInlines.Content.Replace("1990-01-01")

        Dim genderFieldData = CType(formFieldsData("Gender"), FormDropDownData)
        genderFieldData.Items.Add("Select sex")
        genderFieldData.Items.Add("Male")
        genderFieldData.Items.Add("Female")
        genderFieldData.Items.Add("I don't know")
        genderFieldData.HelpText = "Select your gender."
        genderFieldData.StatusText = genderFieldData.HelpText
        genderFieldData.SelectedItemIndex = 0

        Dim marriedFieldData = CType(formFieldsData("Married"), FormCheckBoxData)
        marriedFieldData.HelpText = "Mark as checked if you are married."
        marriedFieldData.StatusText = marriedFieldData.HelpText
        marriedFieldData.DefaultValue = True
        marriedFieldData.Value = True

        Dim salaryFieldData = CType(formFieldsData("Phone"), FormTextData)
        salaryFieldData.TextType = FormTextType.Number
        salaryFieldData.DefaultValue = "555 13-12"
        salaryFieldData.ValueFormat = "(###) ###-####"
        salaryFieldData.HelpText = "Enter your phone number."
        salaryFieldData.StatusText = salaryFieldData.HelpText
        salaryFieldData.Field.ResultInlines.Content.Replace("+1 (800) 111 2233")

        dc.Save("fields-template.pdf", New PdfSaveOptions() With {.PreserveFormFields = True})

        ' Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo("fields-template.pdf") With {.UseShellExecute = True})
    End Sub

End Module