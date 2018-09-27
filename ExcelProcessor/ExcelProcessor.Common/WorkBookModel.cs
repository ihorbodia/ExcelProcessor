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
            workBookFile.MissingCellPolicy = MissingCellPolicy.CREATE_NULL_AS_BLANK;
            fileInfo = new FileInfo(excelFilePathPath);
        }
    }
}
