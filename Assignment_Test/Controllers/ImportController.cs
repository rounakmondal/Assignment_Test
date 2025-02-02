using Assignment_Test.Models;
using System.Collections.Generic;
using System.Web.Mvc;
using System.Net.Http;
using System.Text;
using Newtonsoft.Json;
using System.Threading.Tasks;
using System.Linq;
using System.Web;
using System.Runtime.Remoting.Messaging;
using System;

namespace Assignment_Test.Controllers
{
    public class ImportController : Controller
    {
        // GET: Upload - Display file upload page
        public ActionResult Upload()
        {
            return View();
        }

        // POST: Upload - Upload Excel file to Web API

        [HttpPost]
        public async Task<ActionResult> Upload(HttpPostedFileBase file)
        {
            try { 
            if (file == null || file.ContentLength == 0)
            {
                ViewBag.Error = "Please select a valid Excel file.";
                return View();
            }

                using (var client = new HttpClient())
                {
                    var content = new MultipartFormDataContent();
                    content.Add(new StreamContent(file.InputStream), "file", file.FileName);

                    var response = await client.PostAsync("https://localhost:44361/api/upload", content);
                    if (response.IsSuccessStatusCode)
                    {
                        var result = await response.Content.ReadAsStringAsync();
                        var json = JsonConvert.DeserializeObject<dynamic>(result);
                        TempData["Headers"] = json.headers;
                        TempData["FileName"] = file.FileName;
                        var header = json.headers;
                        var headerArray = json.headers.ToObject<List<string>>();
                        Console.WriteLine(header);
                        try
                        {
                            var model = new MappingViewModel
                            {
                                ExcelHeaders = headerArray,
                                DatabaseFields = GetDatabaseFields()
                            };
                            return View("Mapping", model);
                        }
                        catch (Exception e)
                        {
                            ViewBag.Error = "Error uploading the file.";
                            return View();
                        }

                    }
                    else
                    {
                        return RedirectToAction("Error", "Home", new { message = "Invalid response from service." });
                    }

                }
            }
            catch (Exception ex)
            {
                ViewBag.Error = $"Error occurred: {ex.Message}";
                return RedirectToAction("Error", "Home");
            }
        }

        // POST: Mapping - Submit the mappings to Web API
        [HttpPost]
        public async Task<ActionResult> Import(Dictionary<string, string> mappings)
        {
            try
            {
                using (var client = new HttpClient())
                {                      
                    string fileName = TempData["FileName"].ToString();
                    var requestData = new ImportRequest
                    {
                        FileName = fileName,
                        Mappings = mappings
                    };

                    var json = JsonConvert.SerializeObject(requestData);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await client.PostAsync("https://localhost:44361/api/import", content);

                    if (response.IsSuccessStatusCode)
                    {
                      
                        return RedirectToAction("Result");
                    }
                    else
                    {

                        TempData["ErrorMessage"] = "Import Error  ";
                        return RedirectToAction("Error", "Home");
                    }
                }
            }
            catch (Exception ex)
            {
               
                ViewBag.Error = $"Error occurred: {ex.Message}";
                return RedirectToAction("Error", "Home");
            }
        }

        // Helper: Get Database Fields (You can modify this as per your model)
        private List<string> GetDatabaseFields()
        {
            return typeof(Customer).GetProperties()
                .Where(p => p.Name != "Id")
                .Select(p => p.Name)
                .ToList();
        }

        //responsible for displaying the result view page
        public ActionResult Result()
        {
            return View();
        }
    }
}
