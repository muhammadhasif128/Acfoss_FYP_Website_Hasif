using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ClosedXML.Excel;
using System.IO;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Acfoss_FYP_Website_Hasif.Pages
{
    public class Excel_To_JSON_ConverterModel : PageModel
    {
        [BindProperty]
        public IFormFile excelFile { get; set; }

        public void OnGet()
        {
        }

        // This handler previews the headers of the Excel file.
        public async Task<JsonResult> OnPostPreviewHeadersAsync()
        {
            var headers = new List<string>();
            var file = Request.Form.Files.GetFile("excelFile");

            if (file != null)
            {
                using (var memoryStream = new MemoryStream())
                {
                    await file.CopyToAsync(memoryStream);
                    using (var workbook = new XLWorkbook(memoryStream))
                    {
                        var worksheet = workbook.Worksheet(1);
                        var firstRow = worksheet.FirstRow();
                        foreach (var cell in firstRow.CellsUsed())
                        {
                            headers.Add(cell.Value.ToString());
                        }
                    }
                }
            }

            return new JsonResult(headers);
        }

        // This handler converts the selected fields to JSON.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> OnPostConvertAsync([FromForm] List<string> selectedFields)
        {
            if (excelFile == null || selectedFields == null || selectedFields.Count == 0)
            {
                ModelState.AddModelError("", "Please upload an Excel file and select fields.");
                return Page();
            }

            var jsonData = new List<Dictionary<string, object>>();

            using (var memoryStream = new MemoryStream())
            {
                await excelFile.CopyToAsync(memoryStream);
                using (var workbook = new XLWorkbook(memoryStream))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header row
                    var headers = worksheet.FirstRow().Cells().Select(c => c.Value.ToString()).ToList();

                    foreach (var row in rows)
                    {
                        var rowDict = new Dictionary<string, object>();
                        for (int i = 0; i < headers.Count; i++)
                        {
                            if (selectedFields.Contains(headers[i]))
                            {
                                var header = headers[i];
                                var value = row.Cell(i + 1).Value;
                                rowDict[header] = value ?? "";
                            }
                        }
                        if (rowDict.Count > 0)
                        {
                            jsonData.Add(rowDict);
                        }
                    }
                }
            }

            // Serialize the jsonData to a JSON string
            var jsonResult = JsonConvert.SerializeObject(jsonData, Formatting.Indented);

            // Save the JSON string to a temporary file and provide a download URL
            var tempFileName = "converted-" + Path.GetRandomFileName() + ".json";
            var tempFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "downloads", tempFileName);

            // Ensure the downloads folder exists
            var fileInfo = new FileInfo(tempFilePath);
            fileInfo.Directory.Create(); // If the directory already exists, this method does nothing

            System.IO.File.WriteAllText(tempFilePath, jsonResult);

            // Return the relative URL to the temporary file for download
            return new JsonResult(new { downloadUrl = Url.Content($"~/downloads/{tempFileName}") });
        }
    }
}
