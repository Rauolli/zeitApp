using Microsoft.Office.Interop.Excel;
using System.Drawing.Drawing2D;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace zeitApp
{
    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            ArgumentNullException.ThrowIfNull(args);

            // ExcelTabelle laden
            //string excelPath = new ExcelFileSelector().OpenExcelWorkbook();

            //if (string.IsNullOrEmpty(excelPath))
            //{
            //    MessageBox.Show("Es wurde keine Datei ausgewählt.");
            //    return;
            //}

            string excelPath = @"F:\D+P_Naumann\Zeitabrechnung_24.xlsx";
            // ExcelHandler-Instanz für das ausgewählte Workbook
            ExcelHandlerSingleton excelHandler = ExcelHandlerSingleton.GetInstance(excelPath);

            // Tabellenblätter-Namen laden
            List<string> worksheetNames = excelHandler.GetWorksheetNames();
            
            // Tabellenblatt auswählen
            int workSheetNo = new WorksheetSelector().SelectExcelWorksheet(worksheetNames);
            if (worksheetNames.Count == 0)
            {
                MessageBox.Show("Kein Tabellenblat ausgewählt.");
                return;
            }
            excelHandler.SelectWorkSheet(workSheetNo);

            
            List<WorkDay> workDays = excelHandler.LoadWorkDaysFromSheet();
            WorkMonth workMonth = new WorkMonth(workDays);
            
            excelHandler.WriteWorkDaysToSheet(workMonth);

            //excelHandler.FormatWorksheet();

            excelHandler.SaveAndClose();          
            
        }       

    }
}