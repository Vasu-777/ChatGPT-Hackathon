
public interface IValidationService
{   
   public string hackathonLogFile { get; set; }
    int GetNumberOfColumns(string inputFilePath, char delimiter); //Code from File Format check - Column Names (Niraj) and File format check - No. of column (Niraj)
   string IsValidExcelFile(string file, string FileName, string FilePath);
    bool WriteLog(string path, string error); //Make code robust - Niraj
    bool CheckColumnNames(string inputFilePath, string[] expectedColumns);  //Code from File Format check - Column Names (Niraj) and File format check - No. of column (Niraj)

    public string ValidateDate(); // Code from Data type -  amount (Rishu), Data type checks - Date format (Manraj) 
    string FindDuplicateRowsFromExcel(); //Find duplicity code - Niraj/Balaji
}


