using System.Collections.Generic;

namespace ExcelProcessor.Common
{ 
    public static class CountryFilesHolder
    {
        public static List<WorkBookModel> countryDocFiles;
        public static readonly object locker = new object();
        static CountryFilesHolder()
        {
            countryDocFiles = new List<WorkBookModel>();
        }
    }
}
