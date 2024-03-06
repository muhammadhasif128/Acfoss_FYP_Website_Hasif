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
    }
}
