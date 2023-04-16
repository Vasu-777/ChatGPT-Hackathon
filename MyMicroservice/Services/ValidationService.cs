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



namespace MyMicroservice
{
    public class ValidationService : IValidationService
    {    
    const int numberOfExectedColumns = 11;        
    const string  Col1 = "ClaimNumber"; 
    const string  Col2 = "ClaimCategory";
    const string Col3 = "EmployeeCode";
    const string Col4 = "ClaimDesc";
    const string Col5 = "ReceiptDate"; 
    const string Col6 = "ClaimedAmount";
    const string Col7 = "SubmissionDate";
    const string Col8 = "ClaimStatus"; 
    const string Col9 = "ApprovedAmount";
    const string Col10 = "ApprovalDate";
    const string Col11 = "Approved By";    
   
    public string[][] strErrors = new string[2][]; int intErrors = 0;    
    public string file1 { get; set; }
    public string strMsg{get;set;}
    public string FilePath1{get;set;}
    public string hackathonLogFile { get; set; }
      
     public string IsValidExcelFile(string fileExtension, string FileName, string FilePath) //IFormFile
      {
        file1 = FileName;
        FilePath1 = FilePath;
        if (fileExtension != ".xlsx")
        {
            bool IsCompleted=WriteLog(FilePath1,"Not a xlsx file");
            return "Not a xlsx file"; // false;
        }
        else
        {
            string[] expectedColumns = { Col1.ToString(), Col2.ToString(), Col3.ToString(), Col4.ToString(), Col5.ToString(), Col6.ToString(), Col7.ToString(), Col8.ToString(), Col9.ToString(), Col10.ToString(), Col11.ToString()};                       
            // Call the GetNumberOfColumns method to get the number of columns.
            int numColumns = GetNumberOfColumns(FileName, ',');  
            string ValidFile =  numColumns!=numberOfExectedColumns?"Invalid number of columns in file":"success";

            if(ValidFile=="success" )
            {
                ValidFile=  numColumns==11 && !CheckColumnNames(FileName, expectedColumns)?"Invalid columns name in file":ValidFile;
                strMsg = ValidFile;
                if(strMsg!="success")
                {
                    bool IsCompleted=WriteLog(FilePath1,ValidFile);
                }
                //Console.WriteLine(strMsg);
            }
            else
            {
                strMsg = ValidFile;
                //Console.WriteLine("Called");
                if(strMsg!="success")
                {
                    bool IsCompleted=WriteLog(FilePath1,ValidFile);
                }
                else
                {
                    bool IsCompleted=WriteLog(FilePath1,ValidFile);
                }
                return strMsg;
            }
            if(strMsg == "success")
            {
                
                if(fileExtension == ".xlsx")
                {
                    strMsg = ValidateDate(); //Code from Data type -  amount (Rishu), Data type checks - Date format (Manraj) 
                     strMsg = FindDuplicateRowsFromExcel();
                }
            }
           
        }
         return strMsg.Length > 0 ? strMsg : "success";
      }


//Niraj/Balaji code on duplicity
    //Stat
    // Returns a list of distinct rows based on selected column values from an Excel file
    public string FindDuplicateRowsFromExcel()
    {
        try
        {
            using (var workbook = new XLWorkbook(file1))
                {
                    var worksheet = workbook.Worksheet(1);
                    var rows = worksheet.RowsUsed();

                    // Use LINQ to group rows by their values
                    var groups = rows.GroupBy(r => String.Join(r.Cell(1).Value.ToString(),  r.Cell(2).Value.ToString(),r.Cell(3).Value.ToString(), r.Cell(4).Value.ToString(), r.Cell(5).Value.ToString(), r.Cell(6).Value.ToString(), r.Cell(7).Value.ToString(), r.Cell(8).Value.ToString(), r.Cell(9).Value.ToString(), r.Cell(10).Value.ToString(), r.Cell(11).Value.ToString()));
                // var groups = rows.GroupBy(new r{ r.Cell(1).Value, r.Cell(2).Value}) ;//, r.Cell(3).Value, r.Cell(4).Value, r.Cell(5).Value, r.Cell(6).Value, r.Cell(7).Value, r.Cell(8).Value, r.Cell(9).Value, r.Cell(10).Value, r.Cell(11).Value});

                    // Check if any group has more than one row
                    var duplicates = groups.Where(g => g.Count() > 1);

                    if (duplicates.Any())
                    {
                        //Console.WriteLine("Duplicate data found:");
                        bool IsCompleted=WriteLog(FilePath1,$"Duplicate data found");
                        foreach (var group in duplicates)
                        {
                            //Console.WriteLine(group.Key);
                            foreach (var row in group)
                            {
                                //Console.WriteLine("  Cell: {0}", row.Cell(1).Address);
                                IsCompleted=WriteLog(FilePath1,$" Cell : {row.Cell(1).Address}");
                            }
                            //return false;
                        }
                    }
                    // else
                    // {
                    //     Console.WriteLine("No duplicate data found.");
                    // }
                }

        return "success";
        // }
        }
        catch (Exception ex)
        {
                // Log or handle the exception appropriately
            //Console.WriteLine("Failed to get distinct rows from Excel file: " + ex.Message);
        
        return $"Failed to get distinct rows from Excel File {ex.Message}";
        }
    }

    //Code from File Format check - Column Names (Niraj) and File format check - No. of column (Niraj)
    //Start
    public  int GetNumberOfColumns(string inputFilePath, char delimiter)
    {
        int columnCount =0;
        // Open the input file for reading
                using (var workbook = new XLWorkbook(inputFilePath))
                {
                    // Get the first worksheet in the workbook
                    var worksheet = workbook.Worksheet(1);
                    var rowData = new List<string>();

                    // Get the range of cells containing data in the worksheet
                var range = worksheet.RangeUsed();
                columnCount = range.ColumnCount();
                    
                }

        // Return the length of the array
        return columnCount;
    }
       
    public  bool CheckColumnNames(string filePath, string[] expectedColumns)
    {
        try
        {
            using (var workbook = new XLWorkbook(filePath)) // Open the input file for reading
            {
                var worksheet = workbook.Worksheet(1); // Get the first worksheet in the workbook
                var actualColumns = new List<string>();

                // Iterate over each column in the worksheet's first row and add the column name to actualColumns
                foreach (var cell in worksheet.FirstRow().CellsUsed())
                {
                    actualColumns.Add(cell.Value.ToString());
                }

                // Compare the actual and expected column names
                if (actualColumns.Count != expectedColumns.Length)
                {
                    return false;
                }

                for (int i = 0; i < actualColumns.Count; i++)
                {
                    if (!actualColumns[i].Trim().Equals(expectedColumns[i].Trim()))
                    {
                        return false;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error checking column names for file {filePath}: {ex.Message}");
            return false;
        }

        return true;
    }    
    //End

     //Make code robust - Niraj
    //Start
    public bool WriteLog(string path, string error)
    { 
        bool IsCompleted=false;
        try
        {      
            string logFile = path;
            Console.WriteLine(path);
            logFile += $"\\HackathonLogReport_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
            Console.WriteLine(logFile);
            using (StreamWriter writer = new StreamWriter(logFile, append: true))
            {
                writer.WriteLine($"Log Time: {DateTime.Now}");
                writer.WriteLine($"Message: {error}");
                
            }
            IsCompleted=true;
            hackathonLogFile = logFile;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"{ex.Message} {ex.InnerException} {ex}");
        }
        return IsCompleted;
    }
    //End

     // Code from Data type -  amount (Rishu), Data type checks - Date format (Manraj) 
    //Start
    public string ValidateDate()
    {
        //Console.WriteLine(file1);
        using (var workbook = new XLWorkbook(file1))
        {
            IXLWorksheet worksheet = workbook.Worksheet(1);
                        var invalidDateCells = from row in worksheet.RowsUsed().Skip(1)
                            let dateCells = new IXLCell[] {row.Cell("E"), row.Cell("G"),row.Cell("J")}
                            let invalidCells = dateCells
                            .Where(cell => !cell.IsEmpty())
                            .Where(cell => !DateTime.TryParseExact(cell.GetString(), "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                            where invalidCells.Any()
                            select new {
                                InvalidCells = invalidCells.Select(cell => cell.Address.ToString())
                            };
                foreach (var invalidDateCell in invalidDateCells)
                {
                  //return $"Invalid date value in cell {string.Join(", ", invalidDateCell.InvalidCells)}";
                  bool IsCompleted=WriteLog(FilePath1,$"Invalid date value in cell {string.Join(", ", invalidDateCell.InvalidCells)}");
                  
                }  

                var invalidAmountCells = from row in worksheet.RowsUsed().Skip(1)
                 let amountCells = new IXLCell[] {row.Cell("F"), row.Cell("I")}
                 let invalidCells = amountCells.Where(cell => !cell.IsEmpty())
                                               .Where(cell => !double.TryParse(cell.GetString(), out _))
                                                where invalidCells.Any()
                            select new {
                                InvalidCells = invalidCells.Select(cell => cell.Address.ToString())
                            };
                      foreach (var invalidAmountCell in invalidAmountCells)
                {
                    //return $"Invalid amount value in cell {string.Join(", ", invalidAmountCell.InvalidCells)}";
                    bool IsCompleted=WriteLog(FilePath1,$"Invalid amount value in cell {string.Join(", ", invalidAmountCell.InvalidCells)}");
                }
            // var worksheet = workbook.Worksheet(1);
            // var rowData = new List<string>();

            // // Get the range of cells containing data in the worksheet
            // var range = worksheet.RangeUsed();

            // int lastRow = worksheet.LastRowUsed().RowNumber();

            // string[] columnsToValidate = new string[] { "E", "G", "J" , "F", "I"};
                
            // double result = 0;
            // foreach (string columnName in columnsToValidate)
            // {
            //     // Get the range of cells for the current column
            //     IXLRange columnCells = worksheet.Range(columnName + "2:" + columnName + lastRow);

                
            //     foreach (IXLCell cell in columnCells.Cells())
            //     {
            //         if(columnName.ToString() == "E" || columnName.ToString() == "G" || columnName.ToString() == "J")
            //         {
            //             // Check if the cell is not empty and its value is not in the correct format
            //             if (!DateTime.TryParseExact(cell.GetString(),"dd-MM-yyyy HH:mm:ss", null, System.Globalization.DateTimeStyles.None, out _))
            //             {
            //                 if(!cell.IsEmpty())
            //                 {
            //                     return $"Wrong date format in cell {cell.Address.ToString()}";
            //                     //strErrors[intErrors] = new string[] {$"Wrong date format in cell {cell.Address.ToString()}"};
            //                     //intErrors++;
            //                 }
            //             }
            //         }

            //         if(columnName.ToString() == "F" || columnName.ToString() == "I")
            //         {
            //             var isDouble = double.TryParse(cell.GetString(),out result);

            //             if(!cell.IsEmpty())
            //             {
            //                 if(!isDouble)
            //                 {
            //                     return $"Wrong Amount format in cell {cell.Address.ToString()}";
            //                     //strErrors[intErrors] = new string[]{ $"Wrong Amount format in cell {cell.Address.ToString()}"};
            //                     //intErrors++;
            //                 }
            //             }
            //         }
            //     }
                

            // }
        }
        // string s = ""; Console.Write(strErrors.GetUpperBound(0));
        //                         for(int i = 0; i < strErrors.GetUpperBound(0); i++)
        //                         {
        //                             for(int j= 0; j < strErrors.GetUpperBound(i); j++)
        //                             {
        //                         s = String.Join("\n", strErrors[i][j]);
        //                      Console.WriteLine(s);
                                
        //                         }
        //                         }
                                
        //return strErrors.Length > 0?s: "success";
        return  "success";
    } 
    //End
    
    }

    
}