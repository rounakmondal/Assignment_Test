using Assignment_Test.Models;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Web;
using System.Web.Http;
using Excel = Microsoft.Office.Interop.Excel;


namespace Assignment_Test.Controllers
{
    [RoutePrefix("api")]
    public class ImportApiController : ApiController
    {
        [HttpPost]
        [Route("upload")]
        public IHttpActionResult Upload(HttpRequestMessage request)
        {
            var file = HttpContext.Current.Request.Files["file"];
            if (file == null || file.ContentLength == 0)
                return BadRequest("Please select a valid Excel file.");

            if (Path.GetExtension(file.FileName).ToLower() != ".xlsx")
                return BadRequest("Only Excel files (.xlsx) are supported.");

            string filePath = Path.Combine(HttpContext.Current.Server.MapPath("~/App_Data"), Path.GetFileName(file.FileName));
            file.SaveAs(filePath);

            List<string> excelHeaders = new List<string>();
            try
            {
                var excelApp = new Application();
                var workbooks = excelApp.Workbooks;
                var workbook = workbooks.Open(filePath);
                var worksheet = (Worksheet)workbook.Sheets[1];
                var range = worksheet.UsedRange;

                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    var header = range.Cells[1, col].Text.ToString();
                    if (!string.IsNullOrEmpty(header))
                        excelHeaders.Add(header);
                }
                workbook.Close(false);
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);

                return Ok(new { headers = excelHeaders, fileName = file.FileName });
            }
            catch (Exception ex)
            {
                return InternalServerError(ex);
            }
        }

        [HttpPost]
        [Route("import")]
        public IHttpActionResult ImportData([FromBody] ImportRequest requestData)
        {
            if (requestData == null || requestData.Mappings == null || !requestData.Mappings.Any())
                return BadRequest("No mappings provided.");         
            var fileName = requestData.FileName;
            var mappings = requestData.Mappings;
            if (string.IsNullOrEmpty(fileName))
              
            return BadRequest("FileName is missing in the mappings.");
            var filePath = Path.Combine(HttpContext.Current.Server.MapPath("~/App_Data"), fileName);
                var excelApp = new Excel.Application();
                var workbooks = excelApp.Workbooks;

            try
            {
                excelApp.Visible = false; 
                var workbook = workbooks.Open(filePath);
                var worksheet = workbook.Sheets[1] as Excel.Worksheet;
                var range = worksheet.UsedRange;

                int lastRow = range.Rows.Count;

                using (var db = new ApplicationDbContext())
                {
                    for (int row = 2; row <= lastRow; row++)
                    {
                        var customer = new Customer();
                        foreach (var mapping in requestData.Mappings)
                        {
                            int col = GetColumnIndexByHeader(range, mapping.Value);
                            if (col > 0)
                            {
                                var value = (range.Cells[row, col] as Excel.Range).Text.ToString();
                                typeof(Customer).GetProperty(mapping.Key)?.SetValue(customer, value);
                            }
                        }
                        db.Customers.Add(customer);
                    }
                    db.SaveChanges();
                }               
               
                workbook.Close(false);
                return Ok("Data imported successfully.");
            }
            catch (FileNotFoundException fileEx)
            {
                  return NotFound();
            }
            catch (Exception ex)
            {
                return InternalServerError(ex);
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);

            }
        }

        private int GetColumnIndexByHeader(Excel.Range range, string headerName)
        {
            for (int col = 1; col <= range.Columns.Count; col++)
            {
                if ((range.Cells[1, col] as Excel.Range).Text.ToString() == headerName)
                {
                    return col;
                }
            }
            return -1;
        }
    }
}
