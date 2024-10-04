using ExcelCheckService.Models;
using Newtonsoft.Json;

namespace ExcelCheckService.Services
{
    public class ScanService
    {
        FileHandler _fileHandler;
        string _patternFilePath;
        string _checkFilePath;
        

        public ScanService()
        {
            _fileHandler = new FileHandler();
        }
        public CompareResponse Scan(IFormFile patternFile, IFormFile checkFile, string settings)
        {
            CompareSettings compareSettings;
            if (settings != null)
            {
                compareSettings = JsonConvert.DeserializeObject<CompareSettings>(settings);
            }
            else 
            { 
                compareSettings = null; 
            }

            SavedFilesInfo savedFiles = _fileHandler.SaveFiles(patternFile, checkFile);
            if(savedFiles.Error != "")
            {
                return new CompareResponse()
                {
                    Status = "Not Successful",
                    ComparedCells = 0,
                    FailedCompareCells = 0,
                    Error = savedFiles.Error
                };
            }

            return ScanFiles(savedFiles, compareSettings);
        }

        private CompareResponse ScanFiles(SavedFilesInfo files, CompareSettings compareSettings)
        {
            List<List<ExcelContent>> templateCells = ExcelHelper.ReadExcelFile(files.TemplateFilePath);
            if(compareSettings != null)
            {
                if (compareSettings.checkOnlyCells != null)
                {
                    templateCells = ChekOnlyCells(templateCells, compareSettings.checkOnlyCells);
                }
                if (compareSettings.ignoreCells != null)
                {
                    templateCells = ExcludeCells(templateCells, compareSettings.ignoreCells);
                }
            }
            

            List<List<ExcelContent>> checkCells = ExcelHelper.ReadExcelFile(files.CheckFilePath);

            CompareResponse response = CompareCells(templateCells, checkCells);

            Directory.Delete(Path.GetDirectoryName(files.TemplateFilePath), true);

            return response;
        }

        private List<List<ExcelContent>> ChekOnlyCells(List<List<ExcelContent>> template, ExcelCheckService.Models.Cell[] cells )
        {
            List<List<ExcelContent>> newTemplate = new List<List<ExcelContent>>();
                    List<ExcelContent> newTemplRow = new List<ExcelContent>();
            foreach (var templRow in template)
            {
                newTemplRow.Clear();

                foreach (var templCell in templRow)
                {
                    foreach (var onlyCheckCell in cells)
                    {
                        if(onlyCheckCell.SheetName == templCell.SheetName &&
                            onlyCheckCell.CellName == templCell.CellChar.ToString() + templCell.CellNumber.ToString())
                        {
                            newTemplRow.Add(templCell);
                            
                        }
                    }
                }
                if(newTemplRow.Count > 0)
                {
                    newTemplate.Add(templRow.ToList());
                }
            }
            

            return newTemplate;
        }
        private List<List<ExcelContent>> ExcludeCells(List<List<ExcelContent>> template, ExcelCheckService.Models.Cell[] cells)
        {
            foreach (var templRow in template)
            {
                foreach (var excludeCell in cells)
                {
                    templRow.RemoveAll(cell => cell.CellChar + cell.CellNumber == excludeCell.CellName &&
                    cell.SheetName == excludeCell.SheetName);                    
                }
            }
            
            return template;
        }

        private CompareResponse CompareCells(List<List<ExcelContent>> templateCells, List<List<ExcelContent>> checkCells)
        {
            long processedCells = 0;
            bool isFind = false;
            List<CompareFail> compareFails = new List<CompareFail>();
            
            compareFails.Clear();

            foreach (var templRow in templateCells)
            {
                foreach (var templCell in templRow)
                {
                    isFind = false;

                    foreach (var checkRow in checkCells)
                    {
                        if(templCell.CellNumber == checkRow.FirstOrDefault().CellNumber)
                        {
                            foreach (var checkCell in checkRow)
                            {
                                if(templCell.CellChar == checkCell.CellChar &&
                                   templCell.CellNumber == checkCell.CellNumber &&
                                   templCell.SheetName == checkCell.SheetName)
                                {
                                    processedCells++;
                                    isFind = true;
                                    if(templCell.CellValue.Trim().ToLower() != checkCell.CellValue.Trim().ToLower())
                                    {
                                        compareFails.Add(new CompareFail()
                                        {
                                            TemplateCellValue = templCell.CellValue,
                                            CheckedCellValue = checkCell.CellValue,
                                            CellNumber = templCell.CellChar + templCell.CellNumber,
                                            SheetName = templCell.SheetName
                                        });
                                        break;
                                    }
                                }
                            }                            
                        }
                    }

                    if (!isFind)
                    {
                        processedCells++;
                        compareFails.Add(new CompareFail()
                        {
                            TemplateCellValue = templCell.CellValue,
                            CheckedCellValue = "NoData",
                            CellNumber = templCell.CellChar + templCell.CellNumber,
                            SheetName = templCell.SheetName
                        });
                    }
                }
            }

            CompareResponse compareResponse = new CompareResponse()
            {
                Status = compareFails.Count()>0 ? "Not Successful" : "Successfully",
                ComparedCells = processedCells,
                FailedCompareCells = compareFails.Count(),
                Fails = compareFails.ToArray()
            };
            return compareResponse;
        }
    }
}
