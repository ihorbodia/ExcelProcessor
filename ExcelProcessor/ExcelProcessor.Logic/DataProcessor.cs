using ExcelProcessor.Common;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.AccessControl;

namespace ExcelProcessor.Logic
{
    public class DataProcessor
    {
        ISheet organisationFileDataSheet;
        String excelFileName;
        IWorkbook excelWorkBook;

        public DataProcessor(string filePath)
        {
            if (filePath.Contains("xlsx#"))
            {
                Console.WriteLine("File Error");
            }
            else
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    excelWorkBook = new XSSFWorkbook(file);
                    organisationFileDataSheet = excelWorkBook.GetSheetAt(0);
                    excelFileName = new FileInfo(filePath).Name;
                }
            }
        }

        public static void SetAccessRule(string directory)
        {
            DirectorySecurity sec = System.IO.Directory.GetAccessControl(directory);
            FileSystemAccessRule accRule = new FileSystemAccessRule(Environment.UserDomainName + "\\" + Environment.UserName, FileSystemRights.FullControl, AccessControlType.Allow);
            sec.AddAccessRule(accRule);
        }

        public void ProceedFiles(object callback)
        {
            String dataFromBColumn = "";
            String nameOfOrganisation = "";
            List<IRow> rowsForDelete = new List<IRow>();
            IRow orgRow = null;

            for (int row = 0; row <= organisationFileDataSheet.LastRowNum; row++) //Loop the records upto filled row  
            {
                if (organisationFileDataSheet.GetRow(row) != null) //null is when the row only contains empty cells   
                {
                    orgRow = organisationFileDataSheet.GetRow(row);

                    if (orgRow != null && !ExcelHelper.IsCellEmpty(orgRow.Cells[2]) && orgRow.RowNum > 0)
                    {
                        String orgCellData = ExcelHelper.GetCellData(orgRow.Cells[2]);
                        foreach (WorkBookModel workBookModel in CountryFilesHolder.countryDocFiles)
                        {
                            ISheet countryFileDataSheet = workBookModel.workBookFile.GetSheetAt(2);
                            for (int countryRow = 1; row <= countryFileDataSheet.LastRowNum; countryRow++)
                            {
                                IRow countryRowData = countryFileDataSheet.GetRow(countryRow);
                                if (countryRowData != null)
                                {
                                    if (ExcelHelper.GetCellData(countryRowData.Cells[0]).Contains(orgCellData) && countryRowData.RowNum > 0)
                                    {
                                        string resvalue = ExcelHelper.GetCellData(countryRowData.Cells[1]);
                                        if (countryRowData.Cells[1].CellType == CellType.Formula)
                                        {
                                            var value = workBookModel.workBookFile.GetCreationHelper().CreateFormulaEvaluator().Evaluate(countryRowData.Cells[1]);
                                            resvalue = value.StringValue;
                                        }
                                        dataFromBColumn = resvalue;
                                        nameOfOrganisation = ExcelHelper.GetCellData(countryRowData.Cells[0]);
                                        updateExcelFile(nameOfOrganisation, dataFromBColumn);
                                        rowsForDelete.Add(orgRow);
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void updateExcelFile(String namaOfOrganization, String dataFromBColumn)
        {
            lock (CountryFilesHolder.locker)
            {
                String countryName = excelFileName.Split(' ')[0];
                if (namaOfOrganization != "" && dataFromBColumn != "")
                {
                    var item = CountryFilesHolder.countryDocFiles.Where(x => x.fileInfo.Name.Contains(countryName)).FirstOrDefault();
                    ISheet countryDocSheet = item.workBookFile.GetSheetAt(1);
                    foreach (IRow row in countryDocSheet)
                    {
                        if (ExcelHelper.GetCellData(row.Cells[0]).Contains(namaOfOrganization) && row.RowNum > 0)
                        {
                            ICell cell = row.Cells[1];
                            cell.SetCellValue(dataFromBColumn);
                            return;
                        }
                    }
                }
            }
        }
    }
}
