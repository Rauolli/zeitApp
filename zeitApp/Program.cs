using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.UserSecrets;


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

            var config = new ConfigurationBuilder()
                .AddUserSecrets<Program>() // Falls du secrets verwenden willst
                .Build();

            string? excelPath = config["Excel:Path"];
            if (string.IsNullOrEmpty(excelPath))
            {
                MessageBox.Show("Der Pfad zur Excel-Datei ist nicht konfiguriert.");
                return;
            }

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