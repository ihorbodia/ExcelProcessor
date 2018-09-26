using ExcelProcessor.Common;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

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

                    if (orgRow != null && !ExcelHelper.IsCellEmpty(orgRow.Cells[2]) && orgRow.RowNum > 0)
                    {
                        try
                        {
                            String orgCellData = ExcelHelper.GetCellData(orgRow.Cells[2]);
                            foreach (WorkBookModel workBookModel in CountryFilesHolder.countryDocFiles)
                            {
                                ISheet countryFileDataSheet = workBookModel.workBookFile.GetSheetAt(2);
                                for (int countryRow = 0; row <= countryFileDataSheet.LastRowNum; row++)
                                {
                                    IRow countryRowData = countryFileDataSheet.GetRow(countryRow);
                                    if (ExcelHelper.GetCellData(countryRowData.Cells[0]).Contains(orgCellData) && countryRowData.RowNum > 0)
                                    {
                                        dataFromBColumn = ExcelHelper.GetCellData(countryRowData.Cells[1]);
                                        nameOfOrganisation = ExcelHelper.GetCellData(countryRowData.Cells[0]);
                                        updateExcelFile(nameOfOrganisation, dataFromBColumn);
                                        rowsForDelete.Add(orgRow);
                                        break;
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                }
            }
        }

        private void updateExcelFile(String namaOfOrganization, String dataFromBColumn)
        {
            lock (CountryFilesHolder.locker)
            {
                if (namaOfOrganization != "" && dataFromBColumn != "")
                {
                    foreach (WorkBookModel workBookModel in CountryFilesHolder.countryDocFiles)
                    {
                        String countryName = excelFileName.Split(' ')[0];
                        if (workBookModel.fileInfo.Name.Contains(countryName))
                        {
                            ISheet countryDocSheet = workBookModel.workBookFile.GetSheetAt(1);
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
    }
}
