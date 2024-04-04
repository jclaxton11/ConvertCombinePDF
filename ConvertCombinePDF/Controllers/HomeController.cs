using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using ConvertCombinePDF.Models;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using Microsoft.AspNetCore.Mvc.ModelBinding.Binders;
using Microsoft.Office.Interop.Word;

namespace ConvertCombinePDF.Controllers;

public class HomeController : Controller
{
    private readonly ILogger<HomeController> _logger;
    private static List<byte[]> UploadedFiles = new List<byte[]>();

    public HomeController(ILogger<HomeController> logger)
    {
        _logger = logger;
    }

    public IActionResult Index()
    {
        return View();
    }

    public IActionResult Privacy()
    {
        return View();
    }

    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }

    [HttpPost]
    public IActionResult Upload()
    {
        FilesToConvert FilePathsToConvert = new FilesToConvert();;
        foreach (var file in Request.Form.Files)
        {
            byte[] fileContent = [];
            using (var memoryStream = new MemoryStream())
            {
                file.CopyToAsync(memoryStream);
                fileContent = memoryStream.ToArray();
            }
            FilePathsToConvert.UploadedFilePaths = file.FileName;

            UploadedFiles.Add(fileContent);
        }
        string JSONresult = JsonConvert.SerializeObject(FilePathsToConvert);
        return Ok(JSONresult);
    }

    [HttpPost]
    public async Task<IActionResult> HandlePDF()
    {
        foreach (var file in UploadedFiles)
        {
            byte[] pdfFileContent = await ConvertToPdf(file);
            return File(pdfFileContent, "application/pdf", "output.pdf");
        }
        return Ok();
    }

    public async Task<byte[]> ConvertToPdf(byte[] docxFileStream)
    {
        Console.WriteLine("No way");
        try
        {
            await using (MemoryStream docxStream = new MemoryStream(docxFileStream))
            {
                using (MemoryStream pdfStream = new MemoryStream())
                {
                    using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(docxStream, false))
                    {
                        if (wordDocument?.MainDocumentPart == null)
                        {
                            throw new InvalidOperationException("The main document part is null.");
                        }
                            var body = wordDocument.MainDocumentPart.Document.Body;
                        if (body == null)
                        {
                            throw new InvalidOperationException("The document body is null.");
                        }
                        Console.WriteLine("Creating Document");
                        await using (PdfWriter writer = new PdfWriter(pdfStream))
                        {
                            using (PdfDocument pdf = new PdfDocument(writer))
                            {
                                iText.Layout.Document iTextDocument = new iText.Layout.Document(pdf);
                                
                                // Extract text from Open XML SDK and add to iText document
                                foreach (var paragraph in body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                                {
                                    Console.WriteLine("Added to PDF Document");
                                    iTextDocument.Add(new iText.Layout.Element.Paragraph(paragraph.InnerText));
                                }
                            }
                        }
                    }

                    return pdfStream.ToArray();
                }
            }   
        }
        catch (Exception ex)
        {
            ViewBag.ErrorMessage = ex.Message + "Not able to convert";
            
        }
    return [];
    }

    class FilesToConvert {
        public string UploadedFilePaths = "";
    }
}
