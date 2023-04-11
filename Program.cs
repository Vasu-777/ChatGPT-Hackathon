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
                        FileInfo Files = new FileInfo(UserFiles);
                        
                        var services = new ServiceCollection();
                        services.AddSingleton<IFileValidator1, FileValidator>();
                        
                        var serviceProvider = services.BuildServiceProvider();

                        var myService = serviceProvider.GetService<IFileValidator1>(); 
                        
                        ExcelController fs = new ExcelController(myService, Files.Extension, Files.FullName);
                        
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
