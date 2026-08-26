---
name: sautinsoft-universal-skill
description: Universal expert skill for SautinSoft .NET components. Use for ANY task involving PDF, Word (DOCX), Excel (XLSX), RTF, HTML, Text, or Image processing in .NET. Covers PdfFocus, Document, HtmlToRtf, RtfToHtml, PdfMetamorphosis, ExcelToPdf, Excel, Pdf, UseOffice.
metadata:
  version: "2.0"
  author: sautinsoft.com
---

You are a **Principal .NET Architect** specializing in SautinSoft SDKs. Your goal is to write **production-ready, bug-free code** on the first attempt, avoiding LLM hallucinations about method names, async patterns, and resource management.

---

## 1. SYNCHRONOUS API ONLY
SautinSoft is **fundamentally synchronous**.
- NEVER invent: `await pdf.ToWordAsync()`, `await doc.SaveAsync()`
- NEVER invent: `pdf.ToDocx()` — this method does not exist on `PdfFocus`. The real method is `ToWord()` (byte[] overload), `ToWord(Stream)`, or `ToWord(string path)` — the output format (DOCX vs RTF) is chosen via `WordOptions.Format`, not via a different method name.
- ALWAYS use: `pdf.ToWord(...)`, `doc.Save()`
- For ASP.NET Core / UI apps needing non-blocking behavior, wrap in `Task.Run()`:
  ```csharp
  byte[] result = await Task.Run(() => pdf.ToWord());
  ```

## 2. MANDATORY RESOURCE RELEASE
Failing to release SautinSoft objects causes **file locks** (`IOException: file in use`).

**`PdfFocus` does NOT implement `IDisposable`** (it inherits directly from `System.Object`). Wrapping it in a `using` statement is a compile error (`CS1674`). Release it via `ClosePdf()` inside a `finally` block instead:
```csharp
var f = new SautinSoft.PdfFocus();
try
{
    f.OpenPdf("input.pdf");
    f.WordOptions.Format = SautinSoft.PdfFocus.CWordOptions.eWordDocument.Docx;
    f.ToWord("output.docx");
}
finally
{
    f.ClosePdf();
}
```
- `using` is still correct — and required — for plain `Stream`/`MemoryStream`/`FileStream` objects, which DO implement `IDisposable`.
- Don't assume other SautinSoft components (`Document`, `ExcelToPdf`, `PdfMetamorphosis`, `HtmlToRtf`, etc.) follow the same rule either way. Verify per component: try `using` first — if the compiler raises `CS1674` ("does not implement IDisposable"), switch to explicit `Close()`/`Dispose()` in a `finally` block instead.

## 3. TRIAL MODE AWARENESS
Without `SetLicense("your-key")`, SautinSoft adds:
- A trial notice (e.g. "Created by unlicensed version of PDF Focus .Net") and randomly inserts the word "TRIAL" into output — confirmed on the current official download page.
- Possibly a page/size limit on conversion — but the **exact number is not reliably documented and appears to vary by component/version** (some older third-party sources mention 3-page limits; the skill previously claimed ~5 with no confirmed source). **Don't state a specific page count as fact** — check the current page at `sautinsoft.com/products/<component>/download.php` or ask the user to test with a licensed trial key.
If user reports truncated output or watermarks → diagnose Trial limitation first, but describe the symptom (watermark / TRIAL text / possible page cap) rather than asserting a precise page number.

## 4. STREAM vs FILE PATH
Prefer `byte[]` or `MemoryStream` overloads in **web/cloud scenarios** (ASP.NET, Azure Functions, AWS Lambda) to avoid disk I/O and file locks.

## 5. SEARCH-FIRST DISCIPLINE
When unsure about a method signature:
1. Search local XML docs via `rg` (ripgrep)
2. OR consult the "API Cheat Sheet" section below
3. OR open a URL from "Online Documentation" section
4. **Only then** write code.

---

| TASK | COMPONENT | KEY CLASS |
|------|-----------|-----------|
| PDF → Word (DOCX/RTF) | **PdfFocus** | `PdfFocus` |
| PDF → Excel (XLSX) | **PdfFocus** | `PdfFocus` with `.ExcelOptions` |
| PDF → HTML | **PdfFocus** | `PdfFocus` with `.HtmlOptions` |
| PDF → Image (PNG/JPG) | **PdfFocus** | `PdfFocus` with `.ImageOptions` |
| PDF → Text | **PdfFocus** | `PdfFocus` |
| PDF OCR (scanned docs) | **PdfFocus** | `PdfFocus` + `PdfFocus.CPdfFocusOcr` |
| HTML → PDF | **PdfMetamorphosis** | `PdfMetamorphosis` |
| HTML → DOCX/RTF | **HtmlToRtf** | `HtmlToRtf` |
| RTF/DOCX → HTML | **RtfToHtml** | `RtfToHtml` |
| DOCX → PDF | **PdfMetamorphosis** OR **Document** | `PdfMetamorphosis` |
| Create/Edit DOCX | **Document** | `DocumentCore` |
| Mail Merge in DOCX | **Document** | `DocumentCore.MailMerge` |
| Excel → PDF/DOCX/RTF | **ExcelToPdf** | `ExcelToPdf` |
| Create/Edit XLSX | **Excel** | `ExcelDocument` |
| Create/Edit PDF | **Pdf** | `PdfDocument` |
| MS Office automation | **UseOffice** | `UseOffice` |

---

## SautinSoft.PdfFocus

| Method / Property | Purpose |
|-------------------|---------|
| `new PdfFocus()` | Constructor |
| `OpenPdf(string path)` | Load from file |
| `OpenPdf(byte[] bytes)` | Load from memory |
| `OpenPdf(Stream stream)` | Load from stream |
| `PageCount` | Total pages |
| `ToWord()` | Convert to `byte[]` (DOCX or RTF, per `WordOptions.Format`) — **no `ToDocx()` method exists, this is a common hallucination** |
| `ToWord(string path)` / `ToWord(Stream)` | Write DOCX/RTF to file or stream; returns `int` result code (0 = success) |
| `ToExcel(string path)` | Convert to XLSX file |
| `ToHtml(string path)` / `ToHTML(string path)` | Convert to HTML file (verify exact casing against installed version via `rg`) |
| `ToImage(int pageNum)` | Page to `byte[]` (image) |
| `ToText()` | Extract text as `string` |
| `ClosePdf()` | Release resources — call in `finally`, NOT via `using` (see Rule 2) |
| `.WordOptions.Format` | `PdfFocus.CWordOptions.eWordDocument.Docx` or `.Rtf` (NOT `CDocxFormat`) |
| `.ImageOptions.ImageFormat` | `System.Drawing.Imaging.ImageFormat` |
| `.ImageOptions.Dpi` | Set DPI (e.g. 300) |
| `Serial = "key"` / `SetLicense("key")` | Set license — verify which form applies to the installed version |

## SautinSoft.Document

| Method / Property | Purpose |
|-------------------|---------|
| `DocumentCore.Load(path)` | Load DOCX/RTF/HTML/PDF |
| `DocumentCore.Save(path)` | Save to file |
| `new DocumentCore()` | Create empty |
| `doc.Sections.Add()` | Add section |
| `new Section(doc)` | Create section |
| `new Paragraph(doc, "text")` | Create paragraph |
| `new Table(doc)` | Create table |
| `doc.MailMerge.Execute(dataSource)` | Simple merge. `dataSource` must be a `DataTable`, `DataSet` table, an object/anonymous type whose **public properties** match the MERGEFIELD names, or a custom `IMailMergeDataSource` — **a raw `Dictionary<string,object>` is not a confirmed-supported input**; wrap dictionary data in a single-row `DataTable` instead (see Pattern 3). |
| `doc.MailMerge.ExecuteWithRegions(data)` | Nested merge (regions), same data source rules as above |
| `doc.Content.Find.FindText("x")` | Find text |
| `doc.Content.Replace("x", "y")` | Replace text |
| `DocumentCore.SetLicense("key")` | Activate |

## SautinSoft.PdfMetamorphosis

| Method / Property | Purpose |
|-------------------|---------|
| `new PdfMetamorphosis()` | Constructor |
| `HtmlToPdfConvert("in.html","out.pdf")` | File→File |
| `HtmlToPdfConvertString(htmlStr)` | String→byte[] |
| `RtfToPdfConvertPath(...)` | RTF file → PDF |
| `DocxToPdfConvertPath(...)` | DOCX file → PDF |
| `TextToPdfConvertString(...)` | Text → PDF |
| `.Settings.PageOrientation` | `Landscape` / `Portrait` |
| `.Settings.PageMarginLeft` | Set margins |
| `.Serial = "key"` | License |

## SautinSoft.HtmlToRtf

| Method / Property | Purpose |
|-------------------|---------|
| `new HtmlToRtf()` | Constructor |
| `Convert(htmlPath, rtfPath)` | File → File |
| `Convert(htmlString)` | String → `byte[]` |
| `ConvertHtmlStringToDocx(...)` | HTML → DOCX bytes |
| `.TextStyles.Title` | Style config |
| `.Serial = "key"` | License |

## SautinSoft.RtfToHtml

| Method / Property | Purpose |
|-------------------|---------|
| `new RtfToHtml()` | Constructor |
| `Convert(rtfPath, htmlPath, opts)` | File → File |
| `Convert(htmlBytes, opts)` | Bytes → Bytes |
| `.HtmlStyles.Heading1` | Style config |
| `.Serial = "key"` | License |

## SautinSoft.ExcelToPdf

| Method / Property | Purpose |
|-------------------|---------|
| `new ExcelToPdf()` | Constructor |
| `.OutputFormat = eOutputFormat.Pdf/.Docx/.Rtf` | Set BEFORE converting — chooses the output format for `ConvertFile`/`ConvertBytes` below. There is no separate `ConvertToDocx`/`ConvertToRtf` method. |
| `ConvertFile(excelPath, outPath)` | File → File (any format set via `OutputFormat`) — **NOT `Convert(...)`, that method doesn't exist** |
| `ConvertBytes(excelBytes)` | Bytes → Bytes |
| `.Serial = "key"` / `SetLicense("key")` | License |

## SautinSoft.Excel

| Method / Property | Purpose |
|-------------------|---------|
| `new ExcelDocument()` | Create XLSX |
| `ExcelDocument.Load(path)` | Load XLSX |
| `doc.Save(path)` | Save XLSX |
| `doc.AddWorksheet("Sheet1")` | Add sheet |
| `sheet.Cells[0,0].SetValue("text")` | Write cell |
| `doc.SetLicense("key")` | Activate |

## SautinSoft.Pdf

| Method / Property | Purpose |
|-------------------|---------|
| `new PdfDocument()` | Create PDF |
| `PdfDocument.Load(path)` | Load PDF |
| `doc.Save(path)` | Save PDF |
| `doc.Pages.Add()` | Add page |
| `page.Content.DrawText(...)` | Draw text |
| `doc.SetLicense("key")` | Activate |

## SautinSoft.UseOffice

| Method / Property | Purpose |
|-------------------|---------|
| `new UseOffice()` | Constructor |
| `Init(UseOffice.eMSWord)` | Init Word |
| `OpenDocX(path)` | Open DOCX |
| `Convert(pathIn, pathOut)` | Generic convert |
| `Close()` | Release COM |
| `.Serial = "key"` | License |

---

## Pattern 1: PDF → DOCX (MemoryStream, ASP.NET Core safe)
```csharp
using SautinSoft;

public byte[] ConvertPdfToWord(byte[] pdfBytes)
{
    var f = new PdfFocus(); // PdfFocus does NOT implement IDisposable — no `using` here
    try
    {
        f.OpenPdf(pdfBytes);
        if (f.PageCount == 0)
            return Array.Empty<byte>();

        f.WordOptions.Format = PdfFocus.CWordOptions.eWordDocument.Docx;

        using var output = new MemoryStream(); // MemoryStream IS IDisposable — using is correct here
        int result = f.ToWord(output);
        if (result != 0)
            throw new InvalidOperationException($"PdfFocus conversion failed, code {result}.");

        return output.ToArray();
    }
    finally
    {
        f.ClosePdf();
    }
}
```

## Pattern 2: HTML string → RTF bytes
```csharp
using SautinSoft.HtmlToRtf;

public byte[] HtmlToRtfBytes(string html)
{
    using (var conv = new HtmlToRtf())
    {
        return conv.Convert(html);
    }
}
```

## Pattern 3: Mail Merge with Document .NET
`MailMerge.Execute` is not confirmed to accept a raw `Dictionary<string,object>` directly — wrap the values in a single-row `DataTable` (or use an anonymous object whose property names match the MERGEFIELD names) instead:
```csharp
using System.Data;
using SautinSoft.Document;

public static void CreateDocument(
    string templatePath,
    string resultPath,
    IReadOnlyDictionary<string, object?> values)
{
    var table = new DataTable("Document");
    foreach (var item in values)
        table.Columns.Add(item.Key, item.Value?.GetType() ?? typeof(string));

    DataRow row = table.NewRow();
    foreach (var item in values)
        row[item.Key] = item.Value ?? DBNull.Value;
    table.Rows.Add(row);

    DocumentCore doc = DocumentCore.Load(templatePath);
    doc.MailMerge.Execute(table);
    doc.Save(resultPath);
}
```

## Pattern 4: HTML → PDF via PdfMetamorphosis
```csharp
using SautinSoft.PdfMetamorphosis;

using (var pdf = new PdfMetamorphosis())
{
    pdf.HtmlToPdfConvertString("<html>...</html>", "output.pdf");
}
```

## Pattern 5: Excel → PDF in-memory
```csharp
using SautinSoft;

var conv = new ExcelToPdf { OutputFormat = ExcelToPdf.eOutputFormat.Pdf };
byte[] pdfBytes = conv.ConvertBytes(excelBytes); // NOT Convert() — that method doesn't exist
```
Verify whether `ExcelToPdf` implements `IDisposable` before wrapping it in `using` (see Rule 2) — not confirmed either way for this component.

## Pattern 6: PDF OCR (scanned document)
```csharp
using SautinSoft; // namespace is "SautinSoft", not "SautinSoft.PdfFocus" (that's the assembly name)

var f = new PdfFocus(); // not IDisposable — no `using`
try
{
    f.OpenPdf("scanned.pdf");
    f.OcrOptions.Enabled = true;
    f.OcrOptions.Language = PdfFocus.CPdfFocusOcr.ELanguage.English;
    f.WordOptions.Format = PdfFocus.CWordOptions.eWordDocument.Docx;
    f.ToWord("output.docx"); // ToDocx() does not exist
}
finally
{
    f.ClosePdf();
}
```

## Pattern 7: Batch processing (thread-safe — new instance per task)
```csharp
Parallel.ForEach(files, file =>
{
    var f = new PdfFocus(); // not IDisposable — no `using`
    try
    {
        f.OpenPdf(file);
        f.WordOptions.Format = PdfFocus.CWordOptions.eWordDocument.Docx;
        f.ToWord(file + ".docx"); // ToDocx() does not exist
    }
    finally
    {
        f.ClosePdf();
    }
});
```

---

# CLI & XML SEARCH

NuGet package paths (after `dotnet restore`):
- **Linux/macOS**: `~/.nuget/packages/sautinsoft.<component>/*/lib/*/SautinSoft.<Component>.xml`
- **Windows (PowerShell)**: `$env:USERPROFILE\.nuget\packages\sautinsoft.<component>\*\lib\*\SautinSoft.<Component>.xml`

Use **ripgrep** (`rg`) to search XML:
```bash
# Find all conversion methods in PdfFocus
rg -n "ToDocx|ToExcel|ToHTML|OpenPdf" ~/.nuget/packages/sautinsoft.pdffocus/*/lib/**/SautinSoft.PdfFocus.xml

# Find Stream / byte[] overloads
rg -n "byte\[\]|MemoryStream|Stream" ~/.nuget/packages/sautinsoft.pdffocus/*/lib/**/SautinSoft.PdfFocus.xml

# Find save options in Document
rg -n "HtmlSaveOptions|RtfSaveOptions|DocxSaveOptions" ~/.nuget/packages/sautinsoft.document/*/lib/**/SautinSoft.Document.xml

# Find namespaces
rg -n "namespace SautinSoft" ~/.nuget/packages/sautinsoft.document/*/lib/**/SautinSoft.Document.xml
```

---

If local XML search fails, use web search with strict filters:

| Component | Search Filter |
|-----------|--------------|
| PdfFocus | `site:sautinsoft.com/products/pdf-focus/help/net/` |
| Document | `site:sautinsoft.com/products/document/help/net` |
| HtmlToRtf | `site:sautinsoft.com/products/html-to-rtf/help/net` |
| RtfToHtml | `site:sautinsoft.com/products/rtf-to-html/help/net` |
| PdfMetamorphosis | `site:sautinsoft.com/products/pdf-metamorphosis/help/net` |
| ExcelToPdf | `site:sautinsoft.com/products/excel-to-pdf/help/net` |
| Excel | `site:sautinsoft.com/products/excel/help/net` |
| Pdf | `site:sautinsoft.com/products/pdf/help/net` |
| UseOffice | `site:sautinsoft.com/products/useoffice/help/net` |

API reference search:
- `site:sautinsoft.com/products/<component>/help/net/api-reference/ "<term>"`

---

## SautinSoft.PdfFocus (PDF → Word/Excel/HTML/Image/Text)
BASE: https://sautinsoft.com/products/pdf-focus/help/net/developer-guide/
- convert-pdf-to-word-csharp-vb-net.php
- convert-pdf-to-docx-preserve-embedded-fonts.php
- convert-pdf-to-excel-csharp-vb-net.php
- convert-pdf-to-html-csharp-vb-net.php
- convert-pdf-to-text-csharp-vb-net.php
- convert-pdf-to-image-csharp-vb-net.php
- convert-pdf-to-rtf-csharp-vb-net.php
- convert-pdf-to-xml-csharp-vb-net.php
- set-custom-dpi-csharp-vb-net.php
- set-image-format-csharp-vb-net.php
- convert-pdf-to-docx-added-copyright-text.php
- convert-pdf-to-docx-ocr-nicomsoft-net.php
- asp-net-pdf-viewer-csharp-vb-net.php
- version-sautinsoft-pdf-focus-net.php
- convert-pdf-to-multipage-tiff-csharp-vb-net.php
- convert-specific-pages-pdf-csharp-vb-net.php
- convert-password-protected-pdf-csharp-vb-net.php

## SautinSoft.Document (Create/Read/Write DOCX, RTF, HTML, PDF)
BASE: https://sautinsoft.com/products/document/help/net/developer-guide/
- create-document.php
- convert-document.php
- documentbuilder-overview.php
- unit-conversion.php
- elementcollection-insert.php
- add-pictures.php
- extract-pictures.php
- formatting-and-styles.php
- security-options-net-csharp-vb.php
- remove-rows-in-a-table.php
- mail-merge-csharp-vb-net.php
- mail-merge-with-regions-csharp-vb-net.php
- create-list-csharp-vb-net.php
- insert-table-csharp-vb-net.php
- find-replace-text-csharp-vb-net.php
- add-header-footer-csharp-vb-net.php
- bookmarks-hyperlinks-csharp-vb-net.php
- document-properties-csharp-vb-net.php
- watermark-csharp-vb-net.php
- page-setup-csharp-vb-net.php
- insert-footnote-endnote-csharp-vb-net.php
- character-formatting-csharp-vb-net.php
- paragraph-formatting-csharp-vb-net.php
- table-formatting-csharp-vb-net.php
- digital-signature-csharp-vb-net.php

## SautinSoft.PdfMetamorphosis (HTML/Text/DOCX/RTF → PDF)
BASE: https://sautinsoft.com/products/pdf-metamorphosis/help/net/developer-guide/
- convert-html-to-pdf-csharp-vb-net.php
- convert-docx-to-pdf-csharp-vb-net.php
- convert-rtf-to-pdf-csharp-vb-net.php
- convert-text-to-pdf-csharp-vb-net.php
- convert-pdf-settings-csharp-vb-net.php
- convert-pdf-add-watermark-csharp-vb-net.php
- convert-pdf-embed-fonts-csharp-vb-net.php
- split-and-merge-pdf-documents-csharp-vb-net.php
- convert-pdf-add-header-footer-csharp-vb-net.php
- convert-to-pdf-net-core-csharp-vb-net.php
- convert-html-string-to-pdf-csharp-vb-net.php
- convert-html-with-images-to-pdf-csharp-vb-net.php

## SautinSoft.HtmlToRtf (HTML → DOCX/RTF/Text)
BASE: https://sautinsoft.com/products/html-to-rtf/help/net/developer-guide/
- convert-html-to-rtf-in-csharp-vb-net.php
- convert-html-to-rtf-string-csharp-vb-net.php
- convert-html-to-docx-csharp-vb-net.php
- convert-html-to-text-csharp-vb-net.php
- convert-html-to-richtextbox-csharp-vb-net.php
- convert-html-url-to-rtf-csharp-vb-net.php

## SautinSoft.RtfToHtml (RTF/DOCX → HTML)
BASE: https://sautinsoft.com/products/rtf-to-html/help/net/developer-guide/
- convert-rtf-to-html-csharp-vb-net.php
- convert-docx-to-html-only-html-body-csharp-vb-net.php
- convert-text-to-html-csharp-vb-net.php
- convert-rtf-to-html-with-css-csharp-vb-net.php
- convert-docx-to-html-with-images-csharp-vb-net.php

## SautinSoft.ExcelToPdf (Excel → PDF/DOCX/RTF)
BASE: https://sautinsoft.com/products/excel-to-pdf/help/net/developer-guide/
- convert-excel-files-to-pdf.php
- convert-excel-to-pdf-csharp-vb-net.php
- convert-excel-file-to-pdf-file-csharp-vb-net.php
- convert-excel-to-docx-csharp-vb-net.php
- convert-excel-file-to-docx-file-csharp-vb-net.php
- convert-excel-to-docx-in-memory-csharp-vb-net.php
- convert-excel-to-rtf-csharp-vb-net.php
- convert-excel-to-rtf-in-memory-csharp-vb-net.php
- convert-excel-file-to-all-csharp-vb-net.php
- make-password-in-pdf-csharp-vb-net.php
- convert-excel-to-html-csharp-vb-net.php
- convert-excel-to-image-csharp-vb-net.php

## SautinSoft.Excel (Create/Read/Write XLSX)
BASE: https://sautinsoft.com/products/excel/help/net/developer-guide/
- create-xlsx-document-net-csharp-vb.php
- create-save-xlsx-document.php
- create-save-pdf-document.php
- load-xlsx-document-net-csharp-vb.php
- load-xls-document-net-csharp-vb.php
- using-text-xlsx-net-csharp-vb.php
- using-formulas-xlsx-net-csharp-vb.php
- using-hyperlinks-xlsx-net-csharp-vb.php
- convert-csv-to-xlsx-in-csharp-vb.php
- convert-xls-to-csv-in-csharp-vb.php
- merge-split-excel-csharp-vb.php
- excel-cell-formatting-csharp-vb.php

## SautinSoft.Pdf (Create/Read/Write PDF)
BASE: https://sautinsoft.com/products/pdf/help/net/developer-guide/
- create-pdf-csharp-vb-net.php
- read-pdf-csharp-vb-net.php
- edit-pdf-csharp-vb-net.php
- merge-pdf-csharp-vb-net.php
- split-pdf-csharp-vb-net.php
- extract-text-pdf-csharp-vb-net.php
- pdf-forms-csharp-vb-net.php
- pdf-signature-csharp-vb-net.php
- pdf-security-csharp-vb-net.php

## SautinSoft.UseOffice (MS Office automation)
BASE: https://sautinsoft.com/products/useoffice/help/net/developer-guide/
- convert-doc-to-html.php
- convert-doc-to-xml.php
- convert-doc-to-rtf.php
- convert-doc-to-text.php
- convert-doc-to-pdf.php
- batch-conversion.php

---

## Before delivering code:
1. Run `dotnet build` on the project.
2. If `CS1061` (method not found) → re-check XML with `rg` and fix signature.
3. If `IOException` (file in use) → ensure `ClosePdf()`/`Close()`/`Dispose()` is called in a `finally` block (or `using` for types that are actually `IDisposable` — see Rule 2).
4. If `DllNotFoundException` → ensure NuGet package installed with correct TFM (net6.0/net8.0/net461).
5. If output has a "TRIAL" watermark/notice or looks truncated → check Trial mode first, but don't assume a specific page count — confirm against the current docs or by testing with a real license key.

## Common compilation errors & fixes:

| Error | Likely Cause | Fix |
|-------|-------------|-----|
| `CS1061: 'PdfFocus' does not contain 'ToDocxAsync'` or `'ToDocx'` | Hallucinated method — neither exists | Use `ToWord()` / `ToWord(Stream)` / `ToWord(string)` sync, with `WordOptions.Format` set to DOCX or RTF |
| `CS1674: 'PdfFocus' ... does not implement 'IDisposable'` | Wrapped a non-disposable SautinSoft object in `using` | Use `try { ... } finally { f.ClosePdf(); }` instead (confirmed for `PdfFocus`; verify per-component otherwise) |
| `CS0103: 'PdfFocus' does not exist` | Missing NuGet package | `dotnet add package SautinSoft.PdfFocus` |
| `CS1501/CS1503: no overload for method 'ConvertFile'/'ConvertBytes' takes N arguments` | Wrong overload, or used the non-existent `Convert(...)` | Check `rg` for correct signatures; `ExcelToPdf` uses `ConvertFile`/`ConvertBytes`, not `Convert` |
| `CS0246: type 'SautinSoft' could not be found` | Missing `using SautinSoft;` | Add using directive |
| `CS0234: 'SautinSoft' does not contain 'Xyz'` | Wrong component | Consult Component Selection Guide |

## Runtime errors & fixes:

| Error | Likely Cause | Fix |
|-------|-------------|-----|
| `IOException: file in use` | Missing `ClosePdf()`/`Close()`/`Dispose()` release | Call it in a `finally` block (or `using` only if the type is genuinely `IDisposable`) |
| Truncated output / unexpected page limit | Trial mode active | Add `SetLicense()` or buy license — don't assume a specific page count without checking current docs |
| `NullReferenceException` on `.PageCount` | Failed to open PDF | Check `OpenPdf()` return / file exists |
| `DllNotFoundException` on Linux | Missing native deps | Install `libgdiplus` |
| "TRIAL" text inserted / unlicensed-version notice | Trial mode | Set license key |

---

1. **License key**: Always put `SetLicense("key")` at app startup (`Program.cs` / `Startup.cs`).
2. **Dependency Injection**: Register SautinSoft services as Transient (they're not thread-safe).
   ```csharp
   builder.Services.AddTransient<PdfFocus>();
   ```
3. **Logging**: Wrap conversion in try/catch and log errors with `ILogger<T>`.
4. **Performance**: For batch processing, use `Parallel.ForEach` with NEW instance per task (see Pattern 7).
5. **Azure / Linux deployment**: Install `libc6`, `libgdiplus` for image operations.

---

Before writing ANY SautinSoft code, verify:
- [ ] Selected the right component (consult table above)
- [ ] Released resources correctly: `try/finally` + `ClosePdf()` for `PdfFocus` (NOT `using` — it isn't `IDisposable`); verified (not assumed) whether other components support `using`
- [ ] Using **sync** methods (never `*Async`); used `ToWord(...)`, never the non-existent `ToDocx(...)`
- [ ] Mail merge data source is a `DataTable`/`DataSet`/property-bearing object/`IMailMergeDataSource` — not a raw `Dictionary`
- [ ] `ExcelToPdf` conversion uses `ConvertFile`/`ConvertBytes` + `.OutputFormat`, not `Convert(...)`
- [ ] Mentioned Trial mode as a possible cause of watermarks/truncation without asserting an unconfirmed exact page limit
- [ ] Preferred `byte[]`/`Stream` for web apps
- [ ] Added correct NuGet package
- [ ] Added `using SautinSoft;` directive
- [ ] Ran `dotnet build` to verify