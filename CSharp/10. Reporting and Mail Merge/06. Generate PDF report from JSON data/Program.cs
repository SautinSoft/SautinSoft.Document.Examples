using System;
using System.IO;
using System.Collections.Generic;
using SautinSoft.Document;
using SautinSoft.Document.Drawing;
using SautinSoft.Document.MailMerging;
using Newtonsoft.Json;

/// <summary>
/// Generates a report in PDF format (PDF/A) based on JSON data and .docx template.
/// </summary>
/// <remarks>
/// See details at: https://www.sautinsoft.com/products/document/help/net/developer-guide/mail-merge-generate-pdf-report-from-json-data-net-csharp-vb.php
/// </remarks>
namespace CatBreedReportApp
{
    class CatBreed
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public string PictUrl { get; set; }
        /// <summary>
        /// Weight in lb. (Fields in template: WeightFrom, WeightTo). Here we are using a tuple.
        /// </summary>
        public (int, int) Weight { get; set; }        
    }
    class Program
    {
        static void Main(string[] args)
        {
            // 1. Get json data
            string json = CreateJsonObject();

            // 2. Show json to Console.
            Console.WriteLine(json);

            // 3. Generate report based on .docx template and json.
            GeneratePdfReport(json);
        }

        public static void GeneratePdfReport(string json)
        {
            // Get data from json.            
            var cats = JsonConvert.DeserializeObject<List<CatBreed>>(json);

            // Load the template document.
            string templatePath = @"..\..\..\cats-template.docx";

            DocumentCore dc = DocumentCore.Load(templatePath);

            // To be able to mail merge from your own data source, it must be wrapped into an object that implements the IMailMergeDataSource interface.
            CustomMailMergeDataSource customDataSource = new CustomMailMergeDataSource(cats);

            // Decorate each cat beed by by appropriate picture.
            // Set picture width to 80 mm, height to Auto.
            dc.MailMerge.FieldMerging += (senderFM, eFM) =>
            {
                // Insert an icon before the product name
                if (eFM.RangeName == "CatBreed" && eFM.FieldName == "PictUrl")
                {
                    eFM.Inlines.Clear();
                    string pictPath = eFM.Value.ToString();
                    Picture pict = new Picture(dc, pictPath);
                    double kWH = 1f;
                    double desiredWidthMm = 80;

                    if (pict.Layout.Size.Width > 0 && pict.Layout.Size.Height > 0)
                        kWH = pict.Layout.Size.Width / pict.Layout.Size.Height;

                    pict.Layout = new InlineLayout(new Size(desiredWidthMm, desiredWidthMm / kWH, LengthUnit.Millimeter));

                    eFM.Inlines.Add(pict);                    
                    eFM.Cancel = false;
                }
            };

            // Execute the mail merge.
            dc.MailMerge.Execute(customDataSource);

            string resultPath = "CatBreeds.pdf";

            // Save the output to file
            PdfSaveOptions so = new PdfSaveOptions()
            {
                Compliance = PdfCompliance.PDF_A1a
            };

            dc.Save(resultPath, so);

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(resultPath) { UseShellExecute = true });
        }
        public static string CreateJsonObject()
        {
            string json = String.Empty;
            List<CatBreed> cats = new List<CatBreed>
            {
                new CatBreed() {Title = "Australian Mist",
                    Description = "The Australian Mist (formerly known as the Spotted Mist) is a breed of cat developed in Australia.",
                    PictUrl = "australian-mist.jpg",
                    Weight = (8, 15)},                    
                new CatBreed() {Title = "Maine Coon",
                    Description = "The Maine Coon is a large domesticated cat breed. It has a distinctive physical appearance and valuable hunting skills.",
                    PictUrl = "maine-coon.png",
                    Weight = (13, 18)},                    
                new CatBreed() {Title = "Scottish Fold",
                    Description = "The original Scottish Fold was a white barn cat named Susie, who was found at a farm near Coupar Angus in Perthshire, Scotland, in 1961.",
                    PictUrl = "scottish-fold.jpg",
                    Weight = (9, 13)},
                new CatBreed() {Title = "Oriental Shorthair",
                    Description = "The Oriental Shorthair is a breed of domestic cat that is developed from and closely related to the Siamese cat.",
                    PictUrl = "oriental-shorthair.jpg",
                    Weight = (8, 12)},                    
                new CatBreed() {Title = "Bengal cat",
                    Description = "The earliest mention of an Asian leopard cat × domestic cross was in 1889, when Harrison Weir wrote of them in Our Cats and ...",
                    PictUrl = "bengal-cat.jpg",
                    Weight = (10, 15)},                    
                new CatBreed() {Title = "Russian Blue",
                    Description = "The Russian Blue is a naturally occurring breed that may have originated in the port of Arkhangelsk in Russia.",
                    PictUrl = "russian-blue.jpg",
                    Weight = (8, 15)},
                new CatBreed() {Title = "Mongrel cat",
                    Description = "A mongrel, mutt or mixed-breed cat is a cat that does not belong to one officially recognized breed, but he's cool and gentle!",
                    PictUrl = "mongrel-cat.jpg",
                    Weight = (8, 16)}                    
            };

            // Generate full path for the cat's pictures.
            string pictDirectory = Path.GetFullPath(@"..\..\..\picts\");
            foreach (var cb in cats)
            {
                cb.PictUrl = Path.Combine(pictDirectory, cb.PictUrl);
            }

            // Make serialization to JSON format.            
            json = JsonConvert.SerializeObject(cats);
            return json;
        }

        /// <summary>
        /// A custom mail merge data source that allows SautinSoft.Document to retrieve data from CatBeeds objects.
        /// </summary>
        public class CustomMailMergeDataSource : IMailMergeDataSource
        {
            private readonly List<CatBreed> _cats;
            private int _recordIndex;

            /// <summary>
            /// The name of the data source. 
            /// </summary>
            public string Name
            {
                get { return "CatBreed"; }
            }

            /// <summary>
            /// SautinSoft.Document calls this method to get a value for every data field.
            /// </summary>
            public bool TryGetValue(string valueName, out object value)
            {
                switch (valueName)
                {
                    case "Title":
                        value = _cats[_recordIndex].Title;
                        return true;
                    case "Description":
                        value = _cats[_recordIndex].Description;
                        return true;
                    case "PictUrl":
                        value = _cats[_recordIndex].PictUrl;
                        return true;
                    case "WeightFrom":
                        value = _cats[_recordIndex].Weight.Item1;
                        return true;
                    case "WeightTo":
                        value = _cats[_recordIndex].Weight.Item2;
                        return true;
                    default:
                        // A field with this name was not found
                        value = null;
                        return false;
                }
            }

            /// <summary>
            /// A standard implementation for moving to a next record in a collection.
            /// </summary>
            public bool MoveNext()
            {
                return (++_recordIndex < _cats.Count);
            }

            public IMailMergeDataSource GetChildDataSource(string sourceName)
            {
                return null;
            }
            public CustomMailMergeDataSource(List<CatBreed> cats)
            {
                _cats = cats;
                // When the data source is initialized, it must be positioned before the first record.
                _recordIndex = -1;
            }
        }
    }
}
