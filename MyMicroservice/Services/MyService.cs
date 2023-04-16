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
using System.IO;
using System.Collections.Generic;
//using Microsoft.SharePointOnline.CSOM;
//using Microsoft.IdentityModel.Clients.ActiveDirectory;
 using System.Net.Http;
 using System.Threading.Tasks;

using Microsoft.SharePoint.ClientSideComponent;
using Newtonsoft.Json;

namespace MyMicroservice
{
    public class MyService : IMyService
    {
//         public async Task<string> connectToSharePoint()
//         {
//             var httpClient = new HttpClient();
//             //string siteURL = "https://incedoin-my.sharepoint.com/:x:/g/personal/balaji_ramamurthy_incedoinc_com/Ebz3C7yqxkJNtvvD5qrlaaUBNkAt1n55q2VEzYtiL8Tgcg?e=4%3A30ZQkj&at=9"; //"https://incedoin-my.sharepoint.com/:x:/g/personal/shambhavi_gupta_incedoinc_com/ET2dw0glIvBEuP3JoM2cswkBeBM1Ah7YIJnf1rE5jqattw?e=aea40b"; // "https://
//             // Task<int> task = (Task<int>)Task.Run(async () => {
                
 
//             //     var response = await httpClient.GetAsync(siteURL);
//             //     Task.Delay(100000);
//             // });
//             string siteURL = "https://incedoin-my.sharepoint.com/:x:/g/personal/balaji_ramamurthy_incedoinc_com/Ebz3C7yqxkJNtvvD5qrlaaUBNkAt1n55q2VEzYtiL8Tgcg?e=4%3A30ZQkj&at=9"; //"https://incedoin-my.sharepoint.com/:x:/g/personal/shambhavi_gupta_incedoinc_com/ET2dw0glIvBEuP3JoM2cswkBeBM1Ah7YIJnf1rE5jqattw?e=aea40b"; // "https://
//             //string siteURL = "https://www.google.com/search?q=what+is+integrated+terminal+in+visual+studio+code&rlz=1C1GCEU_enIN953IN953&oq=&aqs=chrome.0.35i39i362l8.206500294j0j15&sourceid=chrome&ie=UTF-8";
//             //var response =  httpClient.GetAsync(siteURL);
// var response = await httpClient.GetAsync(siteURL);
// await Task.Delay(100000);

// // Task<int> task = Task.Run(() => {
// //             // Simulate a long-running operation
// //             Task.Delay(1000).Wait();
// //             return 42;
// //         });

//         // await task.ContinueWith(t => {
//         //     Console.WriteLine($"Task completed with result {t.Result}");
//         //     return ($"Task completed with result {t.Result}");
//         // });

//              if (response.IsSuccessStatusCode)
//             //f(response.IsCompletedSuccessfully)
//             {
//                 // var content = await response.Content.ReadAsStringAsync();
//                 //var content = response.Result; // response.Content.ReadAsStringAsync();
//                 // Do something with the content
//                 string jsonContent = await response.Content.ReadAsStringAsync();
                
//                 // MyObject myObject = JsonSerializer.Deserialize<MyObject>(jsonContent);
//                 var content = JsonSerializer.Deserialize<Object>(jsonContent);
//                 return  string.Join("", content);
//             }
//             else
//             {
//                 // Handle the error
//                 return $"Error {response.Content} {response.StatusCode}";
//             }
//             // return task.Result;
            
//         }
        public string GetMessage()
        {
            return "Hello, World!";
        }

        public string AggregateExcelData(string _fileFullName, string _filePath, string hackathonLogFile) 
        {
            
            var data = new List<string>();

            try
            {
                //string _fileFullName = "C:\\Users\\Balaji.Ramamurthy\\Downloads\\Hackathon\\Hackathon-UseCases2_Data.xlsx";


                string inputFilePath = _fileFullName; // file;
                //FileInfo Files = new FileInfo(inputFilePath);
                string FilePath = _filePath; // Files.Directory.ToString();
                
              //  Console.WriteLine(FilePath);
                string _hackathonLogFile =  hackathonLogFile;// string.IsNullOrWhiteSpace(ValidationService.hackathonLogFile) == true ? "": _IFileValidator1.hackathonLogFile;

                // Open the input file for reading
            //     using (var fileStream = new FileStream(inputFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            // {
                using (var workbook = new XLWorkbook(inputFilePath))
                {
                    //data.Add($"Duplicity check : {GetDistinctRowsFromExcel(inputFilePath)}");
                    
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
                            .Where(row => row.Cell("H").Value.ToString() == "Approved" && double.TryParse(row.Cell("I").Value.ToString(),out _))
                            .GroupBy(row => row.Cell("B").GetString())
                            .Select(row => new {
                                            GroupName = row.Key,
                                            SumApproved = row.Sum(row => row.Cell("I").GetDouble())
                                        })
                                        .ToList();
                    data.Add("\n*************************************************");
                    data.Add($"\nCategory wise total approved claim amount");
                    foreach (var item in query2)
                    {
                        //Console.WriteLine($"{item.GroupName} : {item.SumApproved} ");
                        data.Add($"\n{item.GroupName} : {item.SumApproved} ");
                    }
                    data.Add("\n*************************************************\n");
                    //End   



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
                    data.Add("\n*************************************************");
                    data.Add($"\nMonthly/Quarterly/Yearly claims submission vs approvals.");
                    data.Add("\nMonthly Claims:");
                    foreach (var item in monthlyClaims)
                    {
                        // Console.WriteLine("{0} {1} Submissions: {2}, Approvals: {3}", 
                        //     item.Month, item.Year, item.Submissions, item.Approvals);
                        data.Add($"\n{item.Month} {item.Year} Submissions: {item.Submissions}, Approvals: {item.Approvals}");
                    }

                    //Console.WriteLine("\nQuarterly Claims:");
                    data.Add("\nQuarterly Claims:");
                    foreach (var item in quarterlyClaims)
                    {
                        // Console.WriteLine("{0} {1} Submissions: {2}, Approvals: {3}", 
                        //     item.Quarter, item.Year, item.Submissions, item.Approvals);
                        data.Add($"\n{item.Quarter} , {item.Year} , Submissions: {item.Submissions}, Approvals: {item.Approvals}");
                    }

                    //Console.WriteLine("\nYearly Claims:");
                    data.Add("\nYearly Claims:");
                    foreach (var item in yearlyClaims)
                    {

                        data.Add($"\n{item.Year} Submissions: {item.Submissions}, Approvals: {item.Approvals} ");
                    }
                     data.Add("\n*************************************************\n");
                    data.Add("\n*************************************************");
                    data.Add($"\nProjected Category wise claims for the next quarter, based on the current trend.");
                    //Projected Category wise claims for the next quarter, based on the current trend.Balaji
                    //Start
                    var monthlyClaims1 = from row in claimsData.AsEnumerable()
                                        group row by new {
                                            Year = row.Field<DateTime>("SubmissionDate").Year,
                                            Quarter = (row.Field<DateTime>("SubmissionDate").Month - 1) / 3 + 1,
                                            Month = row.Field<DateTime>("SubmissionDate").Month,
                                            MonthName = monthNames[row.Field<DateTime>("SubmissionDate").Month - 1],
                                            Category = row.Field<String>("ClaimCategory").ToString()
                                        } into grp
                                        orderby grp.Key.Year, grp.Key.Month
                                        select new {
                                            Year = grp.Key.Year,
                                            Quarter = grp.Key.Quarter,
                                            Month = grp.Key.MonthName,
                                            Category = grp.Key.Category,
                                            
                                            Submissions = grp.Count(),
                                            Approvals = grp.Count(r => r.Field<string>("ClaimStatus") == "Approved"),
                                            
                                           PercentageApplied = grp.Count() 
                                        };
                    var quarterlyClaims1 = from row in monthlyClaims1
                                        group row by new {
                                            Year = row.Year,
                                            Quarter = row.Quarter,
                                            Category = row.Category,
                                           // PercentageApplied = row.PercentageApplied
                                        } into grp
                                        orderby grp.Key.Year, grp.Key.Quarter
                                        select new {
                                            Year = grp.Key.Year,
                                            Quarter = string.Format("Q{0}", grp.Key.Quarter),
                                            Submissions = grp.Sum(r => r.Submissions),
                                            Approvals = grp.Sum(r => r.Approvals),
                                            Categories = grp.Key.Category,
                                            PercentageApplied = grp.Sum(r => r.PercentageApplied) //grp.Key.PercentageApplied / grp.Sum(r => r.Submissions) // grp.Sum(r => r.PercentageApplied)
                                        };
                    // data.Add("*************************************************");

                    data.Add("\nProjected categories:");
                    data.Add("\nCategory wise % claimed:");
                    // foreach (var item in monthlyClaims1)
                    // {
                    //     // Console.WriteLine("{0} {1} Submissions: {2}, Approvals: {3}", 
                    //     //     item.Month, item.Year, item.Submissions, item.Approvals);
                    //     data.Add($" {item.Category} {item.Month} {item.Year} Submissions: {item.Submissions}, Approvals: {item.Approvals}");
                    // }
                    
                    int i = claimsData.Rows.Count; // claimsData.Rows.Count; // quarterlyClaims1.Count();
                    data.Add($"\nTotal number of claims found is : {i}"); 
                    //Console.WriteLine($"{quarterlyClaims1.Count()}");
                    foreach (var item in quarterlyClaims1)
                    {
                        Decimal d = Math.Abs(Convert.ToDecimal( item.PercentageApplied) / Convert.ToDecimal(i));
                        float value = (float) d * 100;
                        //data.Add($" ({d}) {item.Categories} {item.Quarter} , {item.Year} , Submissions: {item.Submissions}, Approvals: {item.Approvals}");
                        data.Add($"\n Submitted : {item.PercentageApplied}  %Applied : ({value}) {item.Categories} {item.Quarter} , {item.Year}  "); //Overall Total: {i}
                    }

                    var quarterlyClaims2 = from row in monthlyClaims1
                                        group row by new {
                                            Year = row.Year,
                                            Quarter = row.Quarter,
                                            Category = row.Category,
                                           // PercentageApplied = row.PercentageApplied
                                        } into grp
                                        orderby grp.Key.Year, grp.Key.Quarter
                                        select new {
                                            Year = grp.Key.Quarter+1 > 4? grp.Key.Year+1: grp.Key.Year,
                                            Quarter = string.Format("Q{0}", grp.Key.Quarter+1 > 4? 1: grp.Key.Quarter+1),
                                            Submissions = grp.Sum(r => r.Submissions),
                                            Approvals = grp.Sum(r => r.Approvals),
                                            Categories = grp.Key.Category,
                                            PercentageApplied = grp.Sum(r => r.PercentageApplied) //grp.Key.PercentageApplied / grp.Sum(r => r.Submissions) // grp.Sum(r => r.PercentageApplied)
                                        };
                    data.Add("\nProjected claims:");
                    data.Add($"\nTotal number of claims found is : {i}"); //Projected Total: {i1}
                    int i1 =  claimsData.Rows.Count; // quarterlyClaims2.Count();
                    data.Add($"\nProjected number of claims : {i1}"); 
                    foreach (var item in quarterlyClaims2)
                    {
                        Decimal d = Math.Abs(Convert.ToDecimal( item.PercentageApplied) / Convert.ToDecimal(i1));
                        float value = (float) d * 100;
                        //data.Add($" ({d}) {item.Categories} {item.Quarter} , {item.Year} , Submissions: {item.Submissions}, Approvals: {item.Approvals}");
                        data.Add($"\n Projected : {item.PercentageApplied}  %Applied : ({value}) {item.Categories} {item.Quarter} , {item.Year} ");
                    }
                    data.Add("\n*************************************************\n");
                    //End
                }


                    data.Add("\n*************************************************");
                    data.Add($"\nTotal pending approvals in the system");
                    data.Add("\nTotal Claimed Status "+lastRow);
                    
                    data.Add($"\nTotal Approved Status: {approvedCount}");
                    data.Add($"\nTotal pending approvals in the system: {totalPendingApprovals}");
                    data.Add("\n*************************************************");
                    
                    // Set the path and file name for your text file
                    string filePath = FilePath; // "C:\\Users\\Balaji.Ramamurthy\\Downloads\\Hackathon\\";
                    filePath += $"\\Log_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                    // Open the file and write some text
                    
                    using (StreamWriter writer = new StreamWriter(filePath)) 
                    {
                        // writer.WriteLine("This is some sample text.");
                        // writer.WriteLine("You can write multiple lines.");
                        // writer.WriteLine("And even use variables or other data types.");
                        writer.WriteLine($"\nHackathon Back End Net Core Log Report Created On {DateTime.Now:dd-MM-yyyy HH:mm:ss} ");
                        writer.WriteLine("\n");
                        writer.WriteLine("\n*************************************************");
                        writer.WriteLine("\n");
                        writer.WriteLine($"\nValidations have been performed. ");
                        if(string.IsNullOrWhiteSpace(_hackathonLogFile) == false)
                        {//
                           writer.WriteLine($"Bad records are logged and the file {_hackathonLogFile} are logged in the path {_filePath.ToString()}");
                           writer.WriteLine($"\nPlease refer to the latest file in terms of date and time stamp\n");
                        }
                        
                         writer.WriteLine("\n*************************************************\n");
                        foreach(var items in data)
                        {
                            
                            writer.WriteLine(items);
                        }
                        // writer.WriteLine("*************************************************");
                    }        

                    //End
                }   
                return (string.Join("", data)) ; //data;
            }
            catch (Exception ex)
            {
            data.Add("Some problem encountered : " + ex.Message);
                return (string.Join("", data)) ; //data;
            }
        }

    }

    
}