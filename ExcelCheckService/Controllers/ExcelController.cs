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
        [ApiExplorerSettings(IgnoreApi = true)]
        public async Task<IActionResult> LoadFiles(
            [FromForm] IFormFile patternFile, 
            [FromForm] IFormFile checkFile, 
            [FromForm] string? settings)
        {
            if (patternFile == null) return BadRequest(new { message = "Template file missing" });
            if (checkFile == null) return BadRequest(new { message = "The file being checked is missing" });


            CompareResponse response = _scanService.Scan(patternFile, checkFile, settings);

            return Ok(response);
        }

    }
}
