// See https://aka.ms/new-console-template for more information
using System;
using System.IO;
using FileController1;
using Microsoft.Extensions.DependencyInjection;
using FileValidation;
namespace   HandlingFiles
{
        public  class HandlingFiles
        {
                        
                public static void Main(string[] args)
                {
                
                        string UserFiles = "C:\\Users\\Balaji.Ramamurthy\\Downloads\\Hackathon\\Hackathon-UseCases2_Data.xlsx";
                        //string UserFiles = "C:\\Users\\Balaji.Ramamurthy\\Downloads\\Hackathon\\Hackathon-UseCases2_Data - Wrong Number Of Columns.xlsx";
                        //string UserFiles = "C:\\Users\\Balaji.Ramamurthy\\Downloads\\Hackathon\\Hackathon-UseCases2_Data - Wrong Column Name.xlsx";
                        //string UserFiles = "C:\\Users\\Balaji.Ramamurthy\\Downloads\\Hackathon\\Hackathon_Test_File_Date_And_Amount.xlsx";
                        FileInfo Files = new FileInfo(UserFiles);
                        string FilePath = Files.Directory.ToString();
                        FilePath += "\\";
                        var services = new ServiceCollection();
                        services.AddSingleton<IFileValidator1, FileValidator>();
                        
                        var serviceProvider = services.BuildServiceProvider();

                        var myService = serviceProvider.GetService<IFileValidator1>(); 
                        
                        ExcelController fs = new ExcelController(myService, Files.Extension, Files.FullName, FilePath);
                        
                        if(fs.IsFileValid == "success")
                        {
                                List<string> fileContents = fs.AggregateExcelData();
                                Console.WriteLine(string.Join("\n",fileContents));
                                
                        }
                        else 
                        {
                                Console.WriteLine(fs.IsFileValid);
                        }
                        
                        
                }
        }
}
