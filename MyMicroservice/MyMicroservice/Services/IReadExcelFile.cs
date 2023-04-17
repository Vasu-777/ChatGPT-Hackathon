using Microsoft.AspNetCore.Mvc;
public interface IReadExcelFile
{
    string ReadExcelFileFromHttp([FromForm] IFormFile file);
}