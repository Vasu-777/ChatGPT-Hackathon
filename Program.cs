// See https://aka.ms/new-console-template for more information
using System;
using System.IO;
using FileController1;

namespace   HandlingFiles
{
        public  class HandlingFiles
        {
                public static void Main(string[] args)
                {
                        Console.WriteLine("Hello, World!");
                        string FileName = "C:\\Users\\vasu.nagori\\Documents\\Projects\\ChatGPT_Hackathon\\Project_Documents\\Hackathon-UseCases2_Data";// Path location of Use_Case File in the Local Machine;  
                        ExcelController fs = new ExcelController();
                        List<string> fileContents = fs.ReadFromExcel(FileName); //var fileContents
                        Console.WriteLine(string.Join("\t",fileContents));
                }
        }
}
