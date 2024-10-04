using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Concurrent;
using ExcelCheckService.Models;

namespace ExcelCheckService.Services
{
    public class ExcelHelper
    {
        

        static ConcurrentDictionary<string, List<List<ExcelContent>>> ReadExcelFileData = new ConcurrentDictionary<string, List<List<ExcelContent>>>();

        public static List<List<ExcelContent>> ReadExcelFile(string fileName, bool UseCahe = false)
        {
            var Result = new List<List<ExcelContent>>();

            if (!fileName.ToLower().EndsWith(".xlsx")) return Result;
            if (!System.IO.File.Exists(fileName)) return Result;


            if (UseCahe)
            {
                if (ReadExcelFileData.TryGetValue(fileName, out Result))
                {
                    return Result;
                }
                else
                {
                    Result = new List<List<ExcelContent>>();
                }
            }


            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    var workbook = workbookPart.Workbook;
                    var ListSheets = workbook.Sheets.ToList();

                    foreach (var rec in ListSheets)
                    {
                        string SheetName = rec.GetAttribute("name", "").Value;
                        WorksheetPart worksheetPartFind = GetWorksheetPartByName(doc, SheetName);

                        foreach (SharedStringTablePart sstpart in workbookPart.GetPartsOfType<SharedStringTablePart>())
                        {
                            SharedStringTable sst = sstpart.SharedStringTable;

                            var WorksheetParts = workbookPart.WorksheetParts.ToList();


                            foreach (WorksheetPart worksheetPart in WorksheetParts)
                            {
                                if (worksheetPartFind != worksheetPart) continue;

                                Worksheet sheet = worksheetPart.Worksheet;
                                var rows = sheet.Descendants<Row>();

                                

                                foreach (Row row in rows)
                                {
                                    List<ExcelContent> newRow = new List<ExcelContent>();
                                    foreach (DocumentFormat.OpenXml.Spreadsheet.Cell c in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
                                    {
                                        if (c.CellReference != null)
                                        {
                                            string cellReference = c.CellReference.Value;
                                            string cellChar = new string(cellReference.TakeWhile(char.IsLetter).ToArray());
                                            string cellNumber = new string(cellReference.SkipWhile(char.IsLetter).ToArray());

                                            string cellValue = string.Empty;

                                            
                                            if (c.DataType != null && c.DataType == CellValues.SharedString)
                                            {
                                                int ssid = int.Parse(c.CellValue.Text);
                                                cellValue = sst.ChildElements[ssid].InnerText;
                                            }
                                            else if (c.CellValue != null)
                                            {
                                                cellValue = c.CellValue.Text;
                                            }

                                            
                                            newRow.Add(new ExcelContent()
                                            {
                                                CellChar = cellChar,
                                                CellNumber = cellNumber,
                                                SheetName = SheetName,
                                                CellValue = cellValue
                                            });
                                        }
                                    }

                                    if(newRow.Count > 0)
                                    {
                                        Result.Add(newRow);
                                    }

                                }
                            }
                        }
                    }
                }
            }

            if (UseCahe)
            {
                ReadExcelFileData.TryAdd(fileName, Result);
            }

            return Result;
        }


        private static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            IEnumerable<Sheet> sheets =
               document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
               Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets?.Count() == 0)
            {
                // The specified worksheet does not exist.
                return null;
            }

            string relationshipId = sheets?.First().Id.Value;

            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);

            return worksheetPart;
        }


    }
}
