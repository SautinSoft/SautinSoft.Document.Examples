using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using DocumentCoreMvc.Models;
using SautinSoft.Document;

namespace DocumentCoreMvc.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment environment;

        public HomeController(IWebHostEnvironment environment)
        {
            this.environment = environment;
        }

        public IActionResult Index()
        {
            return View(new InvoiceModel());
        }

        public FileStreamResult Download(InvoiceModel model)
        {
            // Load template document.
            var path = Path.Combine(this.environment.ContentRootPath, "InvoiceWithFields.docx");
            var document = DocumentCore.Load(path);

            // Execute mail merge process.
            document.MailMerge.Execute(model);

            // Save document in specified file format.
            var stream = new MemoryStream();
            document.Save(stream, model.Options);

            // Set the stream position to the beginning of the file.
            //fileStream.Seek(0, SeekOrigin.Begin);

            stream.Seek(0, 0);
            // Download file.
            return File(stream, model.Options.ContentType, $"OutputFromView.{model.Format.ToLower()}");
           
        }

       // [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel() { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}

namespace DocumentCoreMvc.Models
{
    public class InvoiceModel
    {
        public int Number { get; set; } = 1;
        public DateTime Date { get; set; } = DateTime.Today;
        public string Company { get; set; } = "Springfield Nuclear Power Plant (in Sector 7-G)";
        public string Address { get; set; } = "742 Evergreen Terrace, Springfield, United States";
        public string Name { get; set; } = "Homer Simpson";
        public string Format { get; set; } = "DOCX";
        public SaveOptions Options => this.FormatMappingDictionary[this.Format];
        public IDictionary<string, SaveOptions> FormatMappingDictionary => new Dictionary<string, SaveOptions>()
        {
            ["DOCX"] = new DocxSaveOptions(),
            ["HTML"] = new HtmlFixedSaveOptions { EmbedImages =true },
            ["RTF"] = new RtfSaveOptions(),
            ["TXT"] = new TxtSaveOptions(),
            ["PDF"] = new PdfSaveOptions()
        };
    }
}
