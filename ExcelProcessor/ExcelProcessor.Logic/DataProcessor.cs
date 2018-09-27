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
        String excelFilePath;
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
                    excelWorkBook.MissingCellPolicy = MissingCellPolicy.CREATE_NULL_AS_BLANK;
                    organisationFileDataSheet = excelWorkBook.GetSheetAt(0);
                    excelFileName = new FileInfo(filePath).Name;
                    excelFilePath = new FileInfo(filePath).FullName;
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

                    if (orgRow != null && !ExcelHelper.IsCellEmpty(orgRow.GetCell(2)) && orgRow.RowNum > 0)
                    {
                        String orgCellData = ExcelHelper.GetCellData(orgRow.GetCell(2));
                        foreach (WorkBookModel workBookModel in CountryFilesHolder.countryDocFiles)
                        {
                            ISheet countryFileDataSheet = workBookModel.workBookFile.GetSheetAt(2);
                            for (int countryRow = 1; countryRow <= countryFileDataSheet.LastRowNum; countryRow++)
                            {
                                IRow countryRowData = countryFileDataSheet.GetRow(countryRow);
                                if (countryRowData != null)
                                {
                                    if (ExcelHelper.GetCellData(countryRowData.GetCell(0)).Contains(orgCellData) && countryRowData.RowNum > 0)
                                    {
                                        dataFromBColumn = ExcelHelper.GetCellData(countryRowData.GetCell(1));
                                        nameOfOrganisation = ExcelHelper.GetCellData(countryRowData.GetCell(0));
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

            foreach (IRow rowToDelete in rowsForDelete)
            {
                var row = organisationFileDataSheet.GetRow(rowToDelete.RowNum);
                if (row != null)
                {
                    organisationFileDataSheet.RemoveRow(row);
                }
            }

            using (var saveFile = new FileStream(excelFilePath, FileMode.Create, FileAccess.ReadWrite))
            {
                excelWorkBook.Write(saveFile);
                saveFile.Close();
            }
            Console.WriteLine("INFO: File have been processed: " + excelFileName);
        }

        private void updateExcelFile(String nameOfOrganization, String dataFromBColumn)
        {
            lock (CountryFilesHolder.locker)
            {
                String countryName = excelFileName.Split(' ')[0];
                if (nameOfOrganization != "" && dataFromBColumn != "")
                {
                    var item = CountryFilesHolder.countryDocFiles.Where(x => x.fileInfo.Name.Contains(countryName)).FirstOrDefault();
                    ISheet countryDocSheet = item.workBookFile.GetSheetAt(1);
                    foreach (IRow row in countryDocSheet)
                    {
                        if (ExcelHelper.GetCellData(row.GetCell(0)).Contains(nameOfOrganization) && row.RowNum > 0)
                        {
                            ICell cell = row.GetCell(1);
                            cell.SetCellValue(dataFromBColumn);
                            return;
                        }
                    }
                }
            }
        }
    }
}
