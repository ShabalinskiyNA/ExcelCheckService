namespace ExcelCheckService.Models
{
    public class CompareSettings
    {
        public Cell[]? ignoreCells { get; set; }
        public Cell[]? checkOnlyCells { get; set; }
    }
    public class Cell
    {
        public string CellName { get; set; }
        public string SheetName { get; set; }
    }
}