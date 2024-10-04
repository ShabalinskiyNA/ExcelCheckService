namespace ExcelCheckService.Models
{
    public class CompareSettings
    {
        public Cell[] ignoreCells;
        public Cell[] checkOnlyCells;
    }
    public class Cell
    {
        public string CellName { get; set; }
        public string SheetName { get; set; }
    }
}
