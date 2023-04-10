using System;
using System.IO;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.DependencyInjection;
using System.Linq;
using System.Data;

namespace   FileController1
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {

        private readonly IFileValidator1 _excelValidator;
        private string _fileExtension {get; set;}
        private string _fileFullName{get;set;}
        private readonly IFileValidator1 _IFileValidator1;
        public string IsFileValid {get; set;}
        public ExcelController(IFileValidator1 FileValidator1, string fileExtension, string fileFullName)
        {
            _IFileValidator1 = FileValidator1;
            _fileExtension = fileExtension; 
            _fileFullName = fileFullName;
            
            IsFileValid = _IFileValidator1.IsValidExcelFile(_fileExtension, _fileFullName);
        }

        [HttpPost("read")] //Specify this method handles HTTP POST requests from clients
        public List<string> AggregateExcelData() 
        {
            var data = new List<string>();
 
            try
            {

                string inputFilePath = _fileFullName; // file;

                // Open the input file for reading
                using (var workbook = new XLWorkbook(inputFilePath))
                {
                    // Get the first worksheet in the workbook
                    var worksheet = workbook.Worksheet(1);
                    var rowData = new List<string>();

                    // Get the range of cells containing data in the worksheet
                    var range = worksheet.RangeUsed();

                    // Code from Data type -  amount (Rishu) and Data type checks - Date format (Manraj) and Monthly/Quarterly/Yearly claims submission vs approvals (Yanish)
                    //Start
                    // Get the last row in the worksheet
                    int lastRow = worksheet.LastRowUsed().RowNumber();

                    // int lastRow = worksheet.LastRowUsed().RowNumber();
                    var approvedCount = 0;
                    var totalPendingApprovals = 0; //Code from Total pending approvals in the system (Niraj)
                    var claimStatuses = worksheet.Column("H").CellsUsed();

//var data = range.Rows();
var query = worksheet.Rows()
           .Where(row => row.Cell("H").Value.ToString() == "Approved")
           .GroupBy(row => row.Cell("K").GetString())
           .Select(row => new {
                        GroupName = row.Key,
                        SumApproved = row.Sum(row => row.Cell("I").GetDouble())
                    })
                    .ToList();
foreach (var item in query)
{
    Console.WriteLine($"{item.GroupName} : {item.SumApproved} ");
}

                    foreach (var cell in claimStatuses)
                    {
                        // Get the range of cells for the current column
                        if(cell.Value.ToString()=="Approved")
                        approvedCount++; 
                        
                        //Code from Total pending approvals in the system (Niraj)
                        //Start
                        // Get the range of cells for the current column
                        if(cell.Value.ToString()!="Approved")
                        totalPendingApprovals++; 
                        //End

                    }
                        
                    //End

                    // Loop through each row in the range and print the data to the console
                    foreach (var row in range.Rows())
                    {
                        foreach (var cell in row.Cells())
                        {
                            // Get the value of the current cell
                            string cellValue = cell.GetString();
                        
                        }
  
                    }
                    
                    data.Add("\n");
                    data.Add("Total Claimed Status "+lastRow);
                    
                    data.Add("Total Approved Status "+approvedCount);
                    
                    data.Add("Total pending approvals in the system "+totalPendingApprovals);
                    
                    //End
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
