using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using ClosedXML.Excel;
using System.IO;
using Newtonsoft.Json;
using System.Xml;

namespace Acfoss_FYP_Website_Hasif.Pages
{
    public class Excel_To_JSON_ConverterModel : PageModel
    {
        [BindProperty]
        public IFormFile excelFile { get; set; }

        public void OnGet()
        {
        }

        public IActionResult OnPost()
        {
            try
            {
                if (excelFile != null)
                {
                    using (var memoryStream = new MemoryStream())
                    {
                        excelFile.CopyTo(memoryStream);
                        using (var workbook = new XLWorkbook(memoryStream))
                        {
                            var worksheet = workbook.Worksheet(1);
                            var rows = worksheet.RangeUsed().RowsUsed();
                            var headers = rows.First().Cells().Select(c => c.Value.ToString()).ToList();
                            var jsonData = new List<Dictionary<string, object>>();

                            foreach (var row in rows.Skip(1)) // Skip header row
                            {
                                var rowDict = new Dictionary<string, object>();
                                for (int i = 0; i < headers.Count; i++)
                                {
                                    var header = headers[i];
                                    var value = row.Cell(i + 1).Value;
                                    rowDict[header] = value;
                                }
                                jsonData.Add(rowDict);
                            }

                            var jsonResult = JsonConvert.SerializeObject(jsonData, Newtonsoft.Json.Formatting.Indented);

                            return File(System.Text.Encoding.UTF8.GetBytes(jsonResult), "application/json", "Converted.json");
                        }
                    }
                }
                // If the file is null, we return to the page with an error message (or handle it as necessary).
                ModelState.AddModelError("", "Please upload an Excel file.");
                return Page();
            }
            catch (Exception ex)
            {
                // Log the exception, if logging is set up
                // LogException(ex);

                // Inform the user something went wrong
                ModelState.AddModelError("", "An error occurred while processing your request.");
                return Page();
            }
        }

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

        public async Task<IActionResult> OnPostConvertAsync(string[] selectedFields)
        {
            if (excelFile == null || selectedFields == null || selectedFields.Length == 0)
            {
                ModelState.AddModelError("", "Please upload an Excel file and select fields.");
                return Page();
            }

            Console.WriteLine($"Selected Fields: {string.Join(", ", selectedFields)}");

            // Convert selectedFields to lowercase for case-insensitive comparison
            var lowerCaseSelectedFields = selectedFields.Select(sf => sf.ToLower()).ToList();

            var jsonData = new List<Dictionary<string, object>>();

            using (var memoryStream = new MemoryStream())
            {
                await excelFile.CopyToAsync(memoryStream);
                using (var workbook = new XLWorkbook(memoryStream))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skipping header row
                    var headers = worksheet.FirstRow().Cells().Select(c => c.Value.ToString()).ToList();

                    foreach (var row in rows)
                    {
                        var rowDict = new Dictionary<string, object>();
                        for (int i = 0; i < headers.Count; i++)
                        {
                            var headerLowerCase = headers[i].ToLower();
                            if (lowerCaseSelectedFields.Contains(headerLowerCase))
                            {
                                var header = headers[i];
                                var value = row.Cell(i + 1).Value;
                                rowDict.Add(header, value ?? "");
                            }
                        }
                        if (rowDict.Count > 0)
                        {
                            jsonData.Add(rowDict);
                        }
                    }
                }
            }

            // Assuming jsonData is already populated but includes fields not selected by the user
            // Convert selectedFields to a case-insensitive HashSet for efficient lookups
            var selectedFieldSet = new HashSet<string>(selectedFields, StringComparer.OrdinalIgnoreCase);

            for (int j = jsonData.Count - 1; j >= 0; j--)
            {
                var item = jsonData[j];
                var keys = item.Keys.ToList(); // Create a list of keys to iterate over (to avoid modifying the collection while iterating)
                foreach (var key in keys)
                {
                    if (!selectedFieldSet.Contains(key))
                    {
                        item.Remove(key); // Remove the unselected field from the dictionary
                    }
                }
            }

            // Now jsonData only contains data for selected fields and can be serialized to JSON
            var jsonResult = JsonConvert.SerializeObject(jsonData, Newtonsoft.Json.Formatting.Indented);
            return File(System.Text.Encoding.UTF8.GetBytes(jsonResult), "application/json", "Converted.json");
        }
    }
}