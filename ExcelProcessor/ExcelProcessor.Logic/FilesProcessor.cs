using ExcelProcessor.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace ExcelProcessor.Logic
{
    public class FilesProcessor
    {
        string chosenFolderPath;
        string errorMessage;

        public FilesProcessor(string chosenFolderPath)
        {
            this.chosenFolderPath = chosenFolderPath;
        }

        public void InitCountryDocFiles()
        {
            DirectoryInfo diTop = new DirectoryInfo(chosenFolderPath);
            IEnumerable<DirectoryInfo> dirListing = diTop.EnumerateDirectories().Where(s => !s.Name.EndsWith("result"));
            if (dirListing != null)
            {
                foreach (var countryDirectory in dirListing)
                {
                    var docFiles = countryDirectory.EnumerateFiles().Where(s => s.Name.EndsWith("doc.xlsx"));
                    if (docFiles.Count() == 1)
                    {
                        try
                        {
                            CountryFilesHolder.countryDocFiles.Add(new WorkBookModel(docFiles.FirstOrDefault().FullName));
                            Console.WriteLine("INFO: Docfile initialized: " + docFiles.FirstOrDefault().Name);
                        }
                        catch (IOException ex)
                        {
                            Console.WriteLine("ERROR: Problem with docfile: " + docFiles.FirstOrDefault().Name);
                            errorMessage = "ERROR: Problem with country doc file";
                        }
                    }
                }
            }
            else
            {
                errorMessage = "Something wrong with chosen folder";
                Console.WriteLine("ERROR: Something wrong with chosen folder: " + chosenFolderPath);
            }
        }

        public void Run()
        {
            DirectoryInfo diTop = new DirectoryInfo(chosenFolderPath);
            DirectoryInfo resultFolder = diTop.EnumerateDirectories().Where(s => s.Name.EndsWith("result")).FirstOrDefault();

            IEnumerable<FileInfo> filesListing = resultFolder.EnumerateFiles()
                .Where(s => s.Name.Contains("organisation shareholder analysis to do") 
                && !s.Name.Contains("Copie") 
                && !s.Name.Split(new char[0], StringSplitOptions.RemoveEmptyEntries)[0].Contains("old"));

            foreach (var item in filesListing)
            {
                try
                {
                    using (ManualResetEvent resetEvent = new ManualResetEvent(false))
                    {
                        ThreadPool.QueueUserWorkItem(new WaitCallback(x =>
                        {
                            new DataProcessor(item.FullName).ProceedFiles(x);
                            resetEvent.Set();
                        }));
                        resetEvent.WaitOne();
                    }
                }
                catch (Exception ex)
                {
                    errorMessage = "Something wrong with organisation shareholder file: " + item.Name;
                    Console.WriteLine("ERROR: Something wrong with organisation shareholder file: " + item.Name);
                }
            }
            Console.WriteLine("INFO: Organisation shareholders files have been processed");
        }

        public void SaveCountryDocFiles()
        {
            foreach (WorkBookModel countryDocFile in CountryFilesHolder.countryDocFiles)
            {
                countryDocFile.fileStream.Close();
                using (var saveFile = new FileStream(countryDocFile.fileInfoPath, FileMode.Create, FileAccess.Write))
                {
                    countryDocFile.workBookFile.Write(saveFile);
                    saveFile.Close();
                    countryDocFile.workBookFile.Close();
                }
            }
        }
    }
}
