using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace ExcelProcessor.Common
{
    public class WorkBookModel
    {
        public IWorkbook workBookFile;
        public string fileInfoPath;
        public FileStream fileStream;
        public WorkBookModel(string excelFilePathPath)
        {
            FileInfo fp = new FileInfo(excelFilePathPath);
            fileStream = new FileStream(excelFilePathPath, FileMode.Open, FileAccess.Read);
            fileInfoPath = fp.FullName;
            workBookFile = WorkbookFactory.Create(fileStream);
            workBookFile.MissingCellPolicy = MissingCellPolicy.CREATE_NULL_AS_BLANK;
        }
    }
}
