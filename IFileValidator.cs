
using System;
using System.IO;

// using ClosedXML.Excel; 
// using Microsoft.AspNetCore.Hosting;
// using Microsoft.AspNetCore.Http;
// using Microsoft.AspNetCore.Mvc;
// using Microsoft.Extensions.DependencyInjection;

public interface IFileValidator1
{

    string IsValidExcelFile(string file, string FileName, string FilePath); //File extension check (Balaji). IFormFile

    int GetNumberOfColumns(string inputFilePath, char delimiter); //Code from File Format check - Column Names (Niraj) and File format check - No. of column (Niraj)
    bool CheckColumnNames(string inputFilePath, string[] expectedColumns);  //Code from File Format check - Column Names (Niraj) and File format check - No. of column (Niraj)
    
    public string ValidateDate(); // Code from Data type -  amount (Rishu), Data type checks - Date format (Manraj) 
    bool WriteLog(string path, string error); //Make code robust - Niraj
}

