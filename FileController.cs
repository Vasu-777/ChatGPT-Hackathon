using System;
using System.IO;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.DependencyInjection;
using System.Linq;
using System.Data;
using System.Globalization;
using System.Data.OleDb;
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
                    //var approvedCount = 0;
                    //var totalPendingApprovals = 0; //Code from Total pending approvals in the system (Niraj)
                    var claimStatuses = worksheet.Column("H").CellsUsed();

//Vasu code
//Start
//var data = range.Rows();
//Category Wise Total Approved Claim Ammount
var query2 = worksheet.Rows()
           .Where(row => row.Cell("H").Value.ToString() == "Approved")
           .GroupBy(row => row.Cell("B").GetString())
           .Select(row => new {
                        GroupName = row.Key,
                        SumApproved = row.Sum(row => row.Cell("I").GetDouble())
                    })
                    .ToList();
foreach (var item in query2)
{
    //Console.WriteLine($"{item.GroupName} : {item.SumApproved} ");
    data.Add($"{item.GroupName} : {item.SumApproved} ");
}
//End                   

//var data = range.Rows();
// var query = worksheet.Rows()
//            .Where(row => row.Cell("H").Value.ToString() == "Approved")
//            .GroupBy(row => row.Cell("K").GetString())
//            .Select(row => new {
//                         GroupName = row.Key,
//                         SumApproved = row.Sum(row => row.Cell("I").GetDouble())
//                     })
//                     .ToList();
// foreach (var item in query)
// {
//     Console.WriteLine($"{item.GroupName} : {item.SumApproved} ");
// }

// var query1 = worksheet.Rows()
//            .Where(row => row.Cell("H").Value.ToString() == "Approved")
//           // .GroupBy(row => row.Cell("K").GetString())
//            .Select(row => new {
//                         //GroupName = row.Key //,
//                         //SumApproved = row.Sum(row => row.Cell("I").GetDouble())
//                         GroupName = row.Cell("I").GetString()
//                     })
//                     .ToList();
// int count = query1.Count();
// // foreach (var item in query1)
// //{
//     Console.WriteLine($"Approved : {count} ");
// //}

                    // foreach (var cell in claimStatuses)
                    // {
                    //     // Get the range of cells for the current column
                    //     if(cell.Value.ToString()=="Approved")
                    //     approvedCount++; 
                        
                    //     //Code from Total pending approvals in the system (Niraj)
                    //     //Start
                    //     // Get the range of cells for the current column
                    //     if(cell.Value.ToString()!="Approved")
                    //     totalPendingApprovals++; 
                    //     //End

                    // }
                var approvedCount = claimStatuses.Count(cell => cell.Value.ToString() == "Approved");
                var totalPendingApprovals = claimStatuses.Count(cell => cell.Value.ToString() != "Approved");                        

                //Monthly/Quarterly/Yearly claims submission vs approvals (Yanish Code)
                //Start
                //string filePath = @"C:\\Users\\Balaji.Ramamurthy\\Downloads\\Hackathon\\Hackathon-UseCases2_Data.xlsx";
                string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _fileFullName + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';";

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    // Select all rows with non-empty submission dates
                    string query = "SELECT * FROM [Sheet1$] WHERE SubmissionDate IS NOT NULL";

                    OleDbDataAdapter adapter = new OleDbDataAdapter(query, connection);
                    var claimsData = new DataTable();
                    adapter.Fill(claimsData);

                    var monthNames = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.MonthNames;

                    var monthlyClaims = from row in claimsData.AsEnumerable()
                                        group row by new {
                                            Year = row.Field<DateTime>("SubmissionDate").Year,
                                            Quarter = (row.Field<DateTime>("SubmissionDate").Month - 1) / 3 + 1,
                                            Month = row.Field<DateTime>("SubmissionDate").Month,
                                            MonthName = monthNames[row.Field<DateTime>("SubmissionDate").Month - 1]
                                        } into grp
                                        orderby grp.Key.Year, grp.Key.Month
                                        select new {
                                            Year = grp.Key.Year,
                                            Quarter = grp.Key.Quarter,
                                            Month = grp.Key.MonthName,
                                            Submissions = grp.Count(),
                                            Approvals = grp.Count(r => r.Field<string>("ClaimStatus") == "Approved")
                                        };

                    var quarterlyClaims = from row in monthlyClaims
                                        group row by new {
                                            Year = row.Year,
                                            Quarter = row.Quarter
                                        } into grp
                                        orderby grp.Key.Year, grp.Key.Quarter
                                        select new {
                                            Year = grp.Key.Year,
                                            Quarter = string.Format("Q{0}", grp.Key.Quarter),
                                            Submissions = grp.Sum(r => r.Submissions),
                                            Approvals = grp.Sum(r => r.Approvals)
                                        };

                    var yearlyClaims = from row in monthlyClaims
                                    group row by row.Year into grp
                                    orderby grp.Key
                                    select new {
                                        Year = grp.Key,
                                        Submissions = grp.Sum(r => r.Submissions),
                                        Approvals = grp.Sum(r => r.Approvals)
                                    };

                    //Console.WriteLine("Monthly Claims:");
                    data.Add("Monthly Claims:");
                    foreach (var item in monthlyClaims)
                    {
                        // Console.WriteLine("{0} {1} Submissions: {2}, Approvals: {3}", 
                        //     item.Month, item.Year, item.Submissions, item.Approvals);
                        data.Add($"{item.Month} {item.Year} Submissions: {item.Submissions}, Approvals: {item.Approvals}");
                    }

                    //Console.WriteLine("\nQuarterly Claims:");
                    data.Add("\nQuarterly Claims:");
                    foreach (var item in quarterlyClaims)
                    {
                        // Console.WriteLine("{0} {1} Submissions: {2}, Approvals: {3}", 
                        //     item.Quarter, item.Year, item.Submissions, item.Approvals);
                        data.Add($"{item.Quarter} , {item.Year} , Submissions: {item.Submissions}, Approvals: {item.Approvals}");
                    }

                    //Console.WriteLine("\nYearly Claims:");
                    data.Add("\nYearly Claims:");
                    foreach (var item in yearlyClaims)
                    {
                        // Console.WriteLine("{0} Submissions: {1}, Approvals: {2}", 
                        //     item.Year, item.Submissions, item.Approvals);
                        data.Add($"{item.Year} Submissions: {item.Submissions}, Approvals: {item.Approvals} ");
                    }

                }


                //End
                    //End

                    // Loop through each row in the range and print the data to the console
                    // foreach (var row in range.Rows())
                    // {
                    //     foreach (var cell in row.Cells())
                    //     {
                    //         // Get the value of the current cell
                    //         string cellValue = cell.GetString();
                        
                    //     }
  
                    // }
                    
                    //Yanish code - Monthly/Quarterly/Yearly claims submission vs approvals.
                    //Start
                //     var query1 = worksheet.Rows()
                //     .Where(row => row.Cell("H").Value.ToString() == "Approved")
                //    // .GroupBy(row => row.Cell("K").GetString())
                //    .Select(row => new {
                //                 //GroupName = row.Key //,
                //                 //SumApproved = row.Sum(row => row.Cell("I").GetDouble())
                //                 ApprovedDate = row.Cell("I").GetDateTime()
                //             })
                //             .ToList();
                //     var ActualData = data.Where(d => d.ApprovedDate.Month == 1);
                //     // foreach (var item in query1)
                //     //{
                //         Console.WriteLine($"Approved : {count} ");
                //     //}
                    //End

                    data.Add("\n");
                    data.Add("Total Claimed Status "+lastRow);
                    
                    //data.Add("Total Approved Status "+approvedCount);
                    
                    //data.Add("Total pending approvals in the system "+totalPendingApprovals);
                    data.Add($"Total Approved Status: {approvedCount}");
                    data.Add($"Total pending approvals in the system: {totalPendingApprovals}");
                    
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
