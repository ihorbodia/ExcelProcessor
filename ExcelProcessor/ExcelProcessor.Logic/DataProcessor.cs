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
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    excelWorkBook = new XSSFWorkbook(file);
                    excelWorkBook.MissingCellPolicy = MissingCellPolicy.CREATE_NULL_AS_BLANK;
                    organisationFileDataSheet = excelWorkBook.GetSheetAt(0);
                    excelFileName = new FileInfo(filePath).Name;
                    excelFilePath = new FileInfo(filePath).FullName;
                    file.Close();
                }
            }
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
                        string orgCellData = ExcelHelper.GetCellData(orgRow.GetCell(2));
                        foreach (WorkBookModel workBookModel in CountryFilesHolder.countryDocFiles)
                        {
                            ISheet countryFileDataSheet = workBookModel.workBookFile.GetSheetAt(2);
                            if (countryFileDataSheet != null)
                            {
                                for (int countryRow = 1; countryRow <= countryFileDataSheet.LastRowNum; countryRow++)
                                {
                                    IRow countryRowData = countryFileDataSheet.GetRow(countryRow);
                                    if (countryRowData != null)
                                    {
                                        if (ExcelHelper.GetCellData(countryRowData.GetCell(0)).Equals(orgCellData) && countryRowData.RowNum > 0 && !string.IsNullOrEmpty(ExcelHelper.GetCellData(countryRowData.GetCell(1))))
                                        {
                                            dataFromBColumn = ExcelHelper.GetCellData(countryRowData.GetCell(1));
                                            nameOfOrganisation = ExcelHelper.GetCellData(countryRowData.GetCell(0));
                                            updateCountryDocFile(nameOfOrganisation, dataFromBColumn);
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

            foreach (IRow rowToDelete in rowsForDelete)
            {
                var row = organisationFileDataSheet.GetRow(rowToDelete.RowNum);
                if (row != null)
                {
                    organisationFileDataSheet.RemoveRow(row);
                }
            }

            using (var saveFile = new FileStream(excelFilePath, FileMode.Create, FileAccess.Write))
            {
                excelWorkBook.Write(saveFile);
                saveFile.Close();
                excelWorkBook.Close();
            }
            Console.WriteLine("INFO: File have been processed: " + excelFileName);
        }

        private void updateCountryDocFile(string nameOfOrganization, string dataFromBColumn)
        {
            lock (CountryFilesHolder.locker)
            {
                String countryName = excelFileName.Split(' ')[0];
                if (string.IsNullOrWhiteSpace(nameOfOrganization) && string.IsNullOrWhiteSpace(dataFromBColumn))
                {
                    return;
                }
                var countryDocFile = CountryFilesHolder.countryDocFiles.Where(x => x.fileInfoPath.Contains(countryName)).FirstOrDefault();
                if (countryDocFile != null)
                {
                    ISheet countryDocSheet = countryDocFile.workBookFile.GetSheetAt(1);
                    foreach (IRow row in countryDocSheet)
                    {
                        if (ExcelHelper.GetCellData(row.GetCell(0)).Trim().Equals(nameOfOrganization.Trim()) && row.RowNum > 0)
                        {
                            ICell cell = row.GetCell(1);
                            cell.SetCellValue(dataFromBColumn);
                            break;
                        }
                    }
                }
            }
        }
    }
}
