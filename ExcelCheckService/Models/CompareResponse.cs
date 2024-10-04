namespace ExcelCheckService.Models
{
    public class CompareResponse
    {
        public string Status { get; set; }
        public long ComparedCells { get; set; }
        public long FailedCompareCells { get; set; }
        public CompareFail[] Fails { get; set; }
        public string Error { get; set; }
    }
}
