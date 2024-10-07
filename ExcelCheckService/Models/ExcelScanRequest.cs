namespace ExcelCheckService.Models
{
    public class ExcelScanRequest
    {
        public IFormFile patternFile { get; set; }
        public IFormFile checkFile { get; set; }
        public string? compareSettings { get; set; }
    }
}
