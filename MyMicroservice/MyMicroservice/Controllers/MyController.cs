using Microsoft.AspNetCore.Mvc;

[ApiController]
[Route("[controller]")]
public class ExcelController : ControllerBase
{
    private readonly IMyService _myService;

    public ExcelController(IMyService myService)
    {
        _myService = myService;
    }

    //[HttpGet]
    // public ActionResult<string> Get()
    // {
    //     return _myService.GetMessage();
    // }

    [HttpGet]
    public Task<string> Get() // ActionResult<string> Get()
    {
       // return _myService.AggregateExcelData();
        return _myService.connectToSharePoint();
    }
}
