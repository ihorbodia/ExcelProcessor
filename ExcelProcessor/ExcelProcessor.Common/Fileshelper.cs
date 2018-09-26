using System.IO;
using System.Windows.Forms;

namespace ExcelProcessor.Common
{
    public static class FilesHelper
    {
        public static string SelectFolder()
        {
            FolderBrowserDialog openFileDialog = new FolderBrowserDialog();

            string selectedFileName = string.Empty;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFileName = openFileDialog.SelectedPath;
            }
            return selectedFileName;
        }

        
    }
}
