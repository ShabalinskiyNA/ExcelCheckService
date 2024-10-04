namespace ExcelCheckService.Models
{
    public class CompareFail
    {
        public string TemplateCellValue { get; set; }
        public string CheckedCellValue { get; set; }
        public string CellNumber { get; set; }
        public string SheetName { get; set; }
    }
}
