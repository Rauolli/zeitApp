using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace zeitApp
{
    public class ExcelFileSelector
    {
        private readonly string path = @"F:\D+P_Naumann\";
        public string OpenExcelWorkbook()
        {

            string filePath = String.Empty;
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "Excel Files|*.xls;*.xlsx";
            if (dialog.ShowDialog() == true)
            {
                filePath = dialog.FileName;
            }
            return filePath;
        }
    }
}
