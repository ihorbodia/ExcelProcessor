using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcelProcessor.Common
{
    public static class ExcelHelper
    {
		public static void OpenExcelFile()
		{
           
        }

        public static bool IsCellEmpty(ICell cell)
        {
            if (cell == null || cell.CellType == CellType.Blank)
            {
                return true;
            }
            return cell.CellType == CellType.String && string.IsNullOrEmpty(cell.StringCellValue);
        }

        public static string GetCellData(ICell cell)
        {
            if (cell != null)
            {
                if (cell.CellType == CellType.Numeric)
                {
                    return Convert.ToString((cell.NumericCellValue)).Trim();
                }
                else if (cell.CellType == CellType.String)
                {
                    return cell.StringCellValue.Trim();
                }
            }
            return string.Empty;
        }
    }
}
