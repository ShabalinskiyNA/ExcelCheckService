using Microsoft.AspNetCore.Mvc;
using ExcelCheckService.Models;
using ExcelCheckService.Services;


namespace ExcelCheckService.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelController : ControllerBase
    {
        ScanService _scanService;
        public ExcelController()
        {
            _scanService = new ScanService();
        }


        [HttpPost("upload")]
        [Consumes("multipart/form-data")]
        public async Task<IActionResult> LoadFiles([FromForm] ExcelScanRequest excelScanRequest)
        {
            if (excelScanRequest.patternFile == null) return BadRequest(new { message = "Template file missing" });
            if (excelScanRequest.checkFile == null) return BadRequest(new { message = "The file being checked is missing" });


            CompareResponse response = _scanService.Scan(excelScanRequest);

            return Ok(response);
        }

    }
}


