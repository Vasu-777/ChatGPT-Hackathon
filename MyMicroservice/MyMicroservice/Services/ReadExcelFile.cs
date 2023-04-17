using Microsoft.AspNetCore.Mvc;

namespace MyMicroservice
{
    public  class   ReadExcelFile : IReadExcelFile
    {


       [Route("upload-excel")]
       [HttpPost]
        public string ReadExcelFileFromHttp([FromForm] IFormFile file)
        {
            if (file != null && file.Length > 0)
            {
                Console.WriteLine($"Reading Files : {file.FileName}");
                // Process the Excel file and return a response
                return"File uploaded successfully";
            }
            else
            {
                return "BadRequest";
            }
        }
    }
}