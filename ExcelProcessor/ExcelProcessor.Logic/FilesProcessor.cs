using ExcelProcessor.Common;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelProcessor.Logic
{
    public class FilesProcessor
    {
        string chosenFolderPath;
        string organizationFile;
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
                    var docFiles = countryDirectory.EnumerateFiles("doc.xlsx", SearchOption.TopDirectoryOnly);
                    if (docFiles.Count() == 1)
                    {
                        try
                        {
                            CountryFilesHolder.countryDocFiles.Add(new WorkBookModel(docFiles.FirstOrDefault().DirectoryName));
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
                            new DataProcessor(item.DirectoryName).ProceedFiles(x);
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
                errorMessage = "Program interrupted";
                Console.WriteLine("ERROR: Program interrupted");
            }
        }
    }
}
