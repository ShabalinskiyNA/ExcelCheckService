namespace ExcelCheckService.Services
{
    public class FileHandler
    {
        private string _directoryPath;
        public FileHandler()
        {
            _directoryPath = Path.Combine(Directory.GetCurrentDirectory(), "UploadedFiles");
        }
        public SavedFilesInfo SaveFiles(IFormFile templatefile, IFormFile checkfile)
        {
            #region Проверка полученных данных 
            if (templatefile == null || templatefile.Length == 0)
                return new SavedFilesInfo() { Error = "Incorrect template file" };
            if (checkfile == null || checkfile.Length == 0)
                return new SavedFilesInfo() { Error = "The file being checked is not correct" };
            if (!templatefile.FileName.ToLower().EndsWith(".xlsx"))
                return new SavedFilesInfo() { Error = "Invalid template file extension" };
            if (!checkfile.FileName.ToLower().EndsWith(".xlsx"))
                return new SavedFilesInfo() { Error = "Invalid extension of the file being checked" };
            #endregion

            _directoryPath = Path.Combine(_directoryPath, DateTimeOffset.Now.ToUnixTimeMilliseconds().ToString());
            
            string templFilePath = SaveFile(templatefile, _directoryPath);
            string checkFilePath = SaveFile(checkfile, _directoryPath);

            #region 
            if (templFilePath == null)
            {
                return new SavedFilesInfo() { Error = $"Failed to load file - {templatefile.FileName}" };
            }
            if (checkFilePath == null)
            {
                return new SavedFilesInfo() { Error = $"Failed to load file - {checkfile.FileName}" };
            }
            #endregion

            SavedFilesInfo savedFiles = new SavedFilesInfo()
            {
                TemplateFilePath = templFilePath,
                CheckFilePath = checkFilePath,
                Error = ""
            };
            return savedFiles;
        }

        
        private string SaveFile(IFormFile file, string directoryPath)
        {
            if (!Directory.Exists(directoryPath))
            {
                Directory.CreateDirectory(directoryPath);
            }

            string filePath = Path.Combine(directoryPath, Guid.NewGuid().ToString() + file.FileName);

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                file.CopyTo(stream);
            }

            if (File.Exists(filePath)) return filePath;
            else return null;                
        }
    }
    public class SavedFilesInfo
    {
        public string TemplateFilePath { get; set; }
        public string CheckFilePath { get; set; }
        public string Error { get; set; }
    }
}
