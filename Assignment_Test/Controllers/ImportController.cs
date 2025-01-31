using Assignment_Test.Models;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace Assignment_Test.Controllers
{
    public class ImportController : Controller
    {

        //responsible for uploading the Excel file view page
        public ActionResult Upload()
        {
            return View();
        }

        //responsible for uploading the Excel file and reading the headers from the Excel file
        //input: HttpPostedFileBase file
        //output: ActionResult  (save the header's and file name into temp data)

        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            if (file == null || file.ContentLength == 0)
            {
                ViewBag.Error = "Please select a valid Excel file";
                return View();
            }

            if (Path.GetExtension(file.FileName).ToLower() != ".xlsx")
            {
                ViewBag.Error = "Only Excel files (.xlsx) are supported";
                return View();
            }

            string filePath = Path.Combine(Server.MapPath("~/App_Data"), Path.GetFileName(file.FileName));
            file.SaveAs(filePath);

            Application excelApp = new Application();
            Workbook workBook = null;
            Worksheet workSheet = null;
            try
            {
                workBook = excelApp.Workbooks.Open(filePath);
                workSheet = workBook.Sheets[1] as Worksheet;

                if (workSheet == null)
                {
                    throw new Exception("Failed to load worksheet.");
                }
                var excelHeaders = new List<string>();
                Range range = workSheet.UsedRange;
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    var header = (range.Cells[1, col] as Range)?.Value2?.ToString();
                    if (string.IsNullOrEmpty(header)) continue;

                    excelHeaders.Add(header);
                }
                TempData["Headers"] = excelHeaders;
                TempData["FileName"] = file.FileName;
                return RedirectToAction("Mapping");
            }
            catch (Exception ex)
            {
                ViewBag.Error = $"Error occurred: {ex.Message}";
                return View();
            }
            finally
            {
                if (workBook != null) workBook.Close(false);
                if (excelApp != null) excelApp.Quit();
                Marshal.ReleaseComObject(workBook);
                Marshal.ReleaseComObject(excelApp);
               
            }
        }

        //responsilbe for mapping the Excel headers with the database fields view page
        public ActionResult Mapping()
        {
            var headers = TempData["Headers"] as List<string>;
            if (headers == null)
            {
                return RedirectToAction("Upload");
            }

            var model = new MappingViewModel
            {
                ExcelHeaders = headers,
                DatabaseFields = GetDatabaseFields()
            };

            return View(model);
        }
        //responsible for importing the data from the Excel file to the database

        [HttpPost]

        public ActionResult ImportData(Dictionary<string, string> mappings)
        {
            var fileName = TempData["FileName"]?.ToString();
            var filePath = Server.MapPath("~/App_Data/" + fileName);
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workBook = null;
            Excel.Worksheet workSheet = null;

            try
            {
                workBook = excelApp.Workbooks.Open(filePath);
                workSheet = workBook.Sheets[1] as Excel.Worksheet;

                if (workSheet == null)
                {
                    throw new Exception("Failed to load worksheet.");
                }

                using (var db = new ApplicationDbContext())
                {
                    int lastRow = workSheet.UsedRange.Rows.Count;

                    for (int row = 2; row <= lastRow; row++)
                    {
                        var customer = new Customer();

                        foreach (var mapping in mappings)
                        {
                            var excelCol = GetColumnIndex(workSheet, mapping.Value);

                            var dbProperty = mapping.Key;

                            if (excelCol != -1 && !string.IsNullOrEmpty(dbProperty))
                            {
                                var value = (workSheet.Cells[row, excelCol] as Excel.Range)?.Value2?.ToString();
                                typeof(Customer).GetProperty(dbProperty)?.SetValue(customer, value);
                            }
                        }
                        db.Customers.Add(customer);
                    }
                    try
                    {

                    db.SaveChanges();
                    }
                    catch (DbEntityValidationException ex)
                    {
                        var validationErrors = ex.EntityValidationErrors
                            .SelectMany(v => v.ValidationErrors)
                            .Select(e => $"{e.PropertyName}: {e.ErrorMessage}");

                        TempData["Error"] = $"Validation failed: {string.Join("; ", validationErrors.ToString())}";
                        return RedirectToAction("Error", "Home");
                    }
                }
            }
            catch (Exception ex)
            {
                ViewBag.Error = $"Error occurred: {ex.Message}";
                return RedirectToAction("Error","Home");
            }
            finally
            {
                if (workBook != null) workBook.Close(false);
                if (excelApp != null) excelApp.Quit();
                Marshal.ReleaseComObject(workSheet);
                Marshal.ReleaseComObject(workBook);
                Marshal.ReleaseComObject(excelApp);
                System.IO.File.Delete(filePath); 
            }

            return RedirectToAction("Result");
        }

        //responsible for getting the database fields from the Customer model
        private List<string> GetDatabaseFields()
        {
            return typeof(Customer).GetProperties()
                .Where(p => p.Name != "Id")  
                .Select(p => p.Name)
                .ToList();
        }


        //responsible for getting the column index of the Excel file    
        private int GetColumnIndex(Excel.Worksheet worksheet, string header)
        {
            Excel.Range range = worksheet.UsedRange;
            int totalCols = range.Columns.Count;

            for (int col = 1; col <= totalCols; col++)
            {
                var colHeader = (range.Cells[1, col] as Excel.Range)?.Value2?.ToString();
                if (!string.IsNullOrEmpty(colHeader) && colHeader.Equals(header, StringComparison.OrdinalIgnoreCase))
                {
                    return col;
                }
            }
            return -1;
        }

        //responsible for displaying the result view page

        public ActionResult Result()
        {
            return View();
        }   
    }
}
