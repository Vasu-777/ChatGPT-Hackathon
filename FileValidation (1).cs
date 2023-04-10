using System;
using System.IO;

using ClosedXML.Excel; 
// using Microsoft.AspNetCore.Hosting;
// using Microsoft.AspNetCore.Http;
// using Microsoft.AspNetCore.Mvc;
// using Microsoft.Extensions.DependencyInjection;

namespace FileValidation{
public class FileValidator : IFileValidator1
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
    
    public string file1 { get; set; }
    public string strMsg{get;set;}
    public string IsValidExcelFile(string file, string FileName) //IFormFile
    {
        // Add your validation logic here
        // For example, check the file extension and contents to ensure it's a valid Excel file
        // Return true if the file is valid, false otherwise
        
        // if (file1 != ".xlsx")
        file1 = FileName;
        if (file != ".xlsx")
        {
            
            return "Not a xlsx file"; // false;
        }
        else
        {
            
            string[] expectedColumns = { Col1.ToString(), Col2.ToString(), Col3.ToString(), Col4.ToString(), Col5.ToString(), Col6.ToString(), Col7.ToString(), Col8.ToString(), Col9.ToString(), Col10.ToString(), Col11.ToString()};                       
            // Call the GetNumberOfColumns method to get the number of columns.
            int numColumns = GetNumberOfColumns(FileName, ',');  
            string ValidFile =  numColumns!=numberOfExectedColumns?"Invalid number of columns in file":"success";
            
            //if(ValidFile=="success" && numColumns==11 && !myService.CheckColumnNames(UserFiles, expectedColumns))
            if(ValidFile=="success" )
            {
                ValidFile=  numColumns==11 && !CheckColumnNames(FileName, expectedColumns)?"Invalid columns name in file":ValidFile;
                strMsg = ValidFile;
                //Console.WriteLine(strMsg);
            }
            else 
            {
                strMsg = ValidFile;
                //Console.WriteLine("Called");
                return strMsg;
            }

            if(strMsg == "success")
            {
                if(file == ".xlsx")
                {
                    strMsg = ValidateDate(); //Code from Data type -  amount (Rishu), Data type checks - Date format (Manraj) 
                }
            }
            
        }
        
        return strMsg.Length > 0 ? strMsg : "success";
    }

    // Code from Data type -  amount (Rishu), Data type checks - Date format (Manraj) 
    //Start
    public string ValidateDate()
    {
        //Console.WriteLine(file1);
        using (var workbook = new XLWorkbook(file1))
        {
            var worksheet = workbook.Worksheet(1);
            var rowData = new List<string>();

            // Get the range of cells containing data in the worksheet
            var range = worksheet.RangeUsed();

            int lastRow = worksheet.LastRowUsed().RowNumber();

            string[] columnsToValidate = new string[] { "E", "G", "J" , "F", "I"};
                
                double result = 0;
                foreach (string columnName in columnsToValidate)
                {
                    // Get the range of cells for the current column
                    IXLRange columnCells = worksheet.Range(columnName + "2:" + columnName + lastRow);

                  
                    foreach (IXLCell cell in columnCells.Cells())
                    {
                        if(columnName.ToString() == "E" || columnName.ToString() == "G" || columnName.ToString() == "J")
                        {
                            // Check if the cell is not empty and its value is not in the correct format
                            if (!DateTime.TryParseExact(cell.GetString(),"dd-MM-yyyy HH:mm:ss", null, System.Globalization.DateTimeStyles.None, out _))
                            {
                                if(!cell.IsEmpty())
                                {
                                    return $"Wrong date format in cell {cell.Address.ToString()}";
                                }
                            }
                        }

                        if(columnName.ToString() == "F" || columnName.ToString() == "I")
                        {
                            var isDouble = double.TryParse(cell.GetString(),out result);

                            if(!cell.IsEmpty())
                            {
                                if(!isDouble)
                                {
                                    return $"Wrong Amount format in cell {cell.Address.ToString()}";
                                }
                            }
                        }
                    }
                    

                }
        }
        return "success";
    } 
    //End

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
}
}