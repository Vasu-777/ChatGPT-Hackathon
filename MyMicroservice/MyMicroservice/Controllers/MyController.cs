using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("[controller]")]
public class ExcelController : ControllerBase
{
    private readonly IMyService _myService;
    private readonly IValidationService _validationService;
    private readonly IReadExcelFile _readExcelFile;
    public string IsFileValid {get; set;}
     public string _fileFullName{get;set;}
    public string _fileExtension{get;set;}
    public string _filePath{get;set;}
    FileInfo Files;
    public string  _hackathonLogFile  { get; set; }
    
    public ExcelController(IMyService myService,IValidationService validationService, IReadExcelFile readExcelFile)
    {
        
         //_fileFullName = "C:\\Users\\Balaji.Ramamurthy\\Downloads\\Hackathon\\Files\\New Microsoft Excel Worksheet.xlsx";
         
         _fileFullName = "C:\\Users\\Balaji.Ramamurthy\\Downloads\\Hackathon\\Hackathon-UseCases2_Data.xlsx";
         //_fileFullName = "C:\\Users\\Balaji.Ramamurthy\\Downloads\\Hackathon\\Hackathon-UseCases2_Data - Duplicate Data.txt";
        //_fileFullName = "C:\\Users\\Balaji.Ramamurthy\\Downloads\\Hackathon\\Hackathon-UseCases2_Data - Wrong Number Of Columns.xlsx";
        //_fileFullName = "C:\\Users\\Balaji.Ramamurthy\\Downloads\\Hackathon\\Hackathon-UseCases2_Data - Wrong Column Name.xlsx";
        //_fileFullName = "C:\\Users\\Balaji.Ramamurthy\\Downloads\\Hackathon\\Hackathon_Test_File_Date_And_Amount.xlsx";
        //_fileFullName = "C:\\Users\\Balaji.Ramamurthy\\Downloads\\Hackathon\\Hackathon-UseCases2_Data - Duplicate Data.xlsx";
                        
         Files = new FileInfo(_fileFullName);
         _fileExtension = Files.Extension;
         _filePath = Files.Directory.ToString();
         Console.WriteLine($"FIle directory : {_filePath}");
        _myService = myService;
        _validationService = validationService;
        IsFileValid = _validationService.IsValidExcelFile(_fileExtension, _fileFullName, _filePath);
        _hackathonLogFile = _validationService.hackathonLogFile;

        _readExcelFile = readExcelFile;
    }

    //[HttpGet]
    // public ActionResult<string> Get()
    // {
    //     return _myService.GetMessage();
    // }

    [HttpGet]
    public ActionResult<string> Get() // Task<string> Get()
    {
        if(IsFileValid=="success")
        {
            return _myService.AggregateExcelData(_fileFullName, _filePath, _hackathonLogFile);
            //return _myService.connectToSharePoint();
        }
        else 
        {
            return IsFileValid;
        }
    }
    // [HttpGet]
    // public ActionResult<string> Get()
    // {
    //     Console.WriteLine("Hello");
    //     IFormFile formFileFile;
        
    //     using (var stream = System.IO.File.OpenRead($"{_fileFullName}"))
    //     {
    //         formFileFile = new FormFile(stream, 0, stream.Length, null, Path.GetFileName(Files.Name));
    //     }
    //     //IActionResult response = _readExcelFile.ReadExcelFileFromHttp(formFileFile);
    //     //return response; // _readExcelFile.ReadExcelFileFromHttp(formFileFile);
    //    //return Ok();
       
    //    return _readExcelFile.ReadExcelFileFromHttp(formFileFile);
    //    //return msg;
    // }
}
