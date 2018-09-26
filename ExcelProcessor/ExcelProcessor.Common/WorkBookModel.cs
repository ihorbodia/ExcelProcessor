using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace ExcelProcessor.Common
{
    public class WorkBookModel
    {
        public IWorkbook workBookFile;
        public FileInfo fileInfo;
        public WorkBookModel(string excelFilePathPath)
        {
            workBookFile = new XSSFWorkbook(excelFilePathPath);
            fileInfo = new FileInfo(excelFilePathPath);
        }
    }
}
