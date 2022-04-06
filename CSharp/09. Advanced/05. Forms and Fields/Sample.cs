using System;
using System.Globalization;
using System.Text;
using SautinSoft.Document;
using SautinSoft.Document.MailMerging;
using SautinSoft.Document.Tables;

class Sample
{
    static void Main(string[] args)
    {
        FormsAndFields();
    }
	
    /// <summary>
    /// Generate document (PDF) with forms and fields.
    /// </summary>
    /// <remarks>
    /// Details: https://www.sautinsoft.com/products/document/help/net/developer-guide/forms-and-fields.php
    /// </remarks>
    public static void FormsAndFields()
    {
        DocumentCore dc = new DocumentCore();

        string placeHolder = new string('\x2002', 50);

        // Create form fields.
        Field fFullName = new Field(dc, FieldType.FormText, null, placeHolder);
        fFullName.FormData.Name = "FullName";
        fFullName.FormData.Enabled = true;


        Field fBirthData = new Field(dc, FieldType.FormText, null, placeHolder);
        fBirthData.FormData.Name = "BirthDate";

        Field fGender = new Field(dc, FieldType.FormDropDown);
        fGender.FormData.Name = "Gender";

        Field fMarried = new Field(dc, FieldType.FormCheckBox);
        fMarried.FormData.Name = "Married";
        fMarried.FormData.Enabled = true;

        Field fPhone = new Field(dc, FieldType.FormText, null, placeHolder);
        fPhone.FormData.Name = "Phone";

        dc.Sections.Add(new Section(dc,

            new Paragraph(dc,
            new Run(dc, "Full name: "),
            fFullName),

            new Paragraph(dc,
            new Run(dc, "Birth date: "),
            fBirthData),

            new Paragraph(dc,
            new Run(dc, "Gender: "),
            fGender),

            new Paragraph(dc,
            new Run(dc, "Married: "),
            fMarried),

            new Paragraph(dc,
            new Run(dc, "Phone: "),
            fPhone)));


        // Customize form fields.
        var formFieldsData = dc.Content.FormFieldsData;

        var fullNameFieldData = (FormTextData)formFieldsData["FullName"];
        fullNameFieldData.MaximumLength = 50;
        fullNameFieldData.StatusText = fullNameFieldData.HelpText = "Enter your name and surname (trimmed to 50 characters).";
        fullNameFieldData.Field.ResultInlines.Content.Replace("Mister Bean");

        var birthdateFieldData = (FormTextData)formFieldsData["BirthDate"];
        birthdateFieldData.TextType = FormTextType.Date;
        birthdateFieldData.DefaultValue = "1990-01-01";
        birthdateFieldData.ValueFormat = "yyyy-MM-dd";
        birthdateFieldData.StatusText = birthdateFieldData.HelpText =
            "Enter your date of birth.";
        birthdateFieldData.Field.ResultInlines.Content.Replace("1990-01-01");

        var genderFieldData = (FormDropDownData)formFieldsData["Gender"];
        genderFieldData.Items.Add("Select sex");
        genderFieldData.Items.Add("Male");
        genderFieldData.Items.Add("Female");
        genderFieldData.Items.Add("I don't know");
        genderFieldData.StatusText = genderFieldData.HelpText =
            "Select your gender.";
        genderFieldData.SelectedItemIndex = 0;

        var marriedFieldData = (FormCheckBoxData)formFieldsData["Married"];
        marriedFieldData.StatusText = marriedFieldData.HelpText =
                    "Mark as checked if you are married.";
        marriedFieldData.DefaultValue = true;
        marriedFieldData.Value = true;

        var salaryFieldData = (FormTextData)formFieldsData["Phone"];
        salaryFieldData.TextType = FormTextType.Number;
        salaryFieldData.DefaultValue = "555 13-12";
        salaryFieldData.ValueFormat = "(###) ###-####";
        salaryFieldData.StatusText = salaryFieldData.HelpText =
            "Enter your phone number.";
        salaryFieldData.Field.ResultInlines.Content.Replace("+1 (800) 111 2233");

        dc.Save(@"fields-template.pdf", new PdfSaveOptions() {PreserveFormFields=true });

        // Open the result for demonstration purposes.
        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(@"fields-template.pdf") { UseShellExecute = true });
    }
}