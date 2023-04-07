using System;
using System.IO;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace   FileController1
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {

        [HttpPost("read")] //Specify this method handles HTTP POST requests from clients
        public List<string> ReadFromExcel(string file) 
        {
            var data = new List<string>();
            try
            {
                if (Path.GetExtension(file) != ".xlsx")
                {
                    data.Add("The selected file is not an excel file");
                    return data;
                }

                string inputFilePath = file;

            // Open the input file for reading
            using (var workbook = new XLWorkbook(inputFilePath))
            {
                // Get the first worksheet in the workbook
                var worksheet = workbook.Worksheet(1);
                var rowData = new List<string>();

                // Get the range of cells containing data in the worksheet
                var range = worksheet.RangeUsed();

                // Loop through each row in the range and print the data to the console
                foreach (var row in range.Rows())
                {
                    foreach (var cell in row.Cells())
                    {
                        // Get the value of the current cell
                        string cellValue = cell.GetString();
                        rowData.Add(cellValue); data.Add(cellValue);
                    }
                    rowData.Add("\n");
                    data.Add("\n");
                }
            }
            return data;
            }
            catch (Exception ex)
            {
            data.Add(ex.Message);
                return data;
            }
        }
    }
}
