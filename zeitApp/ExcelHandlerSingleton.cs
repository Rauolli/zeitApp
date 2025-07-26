using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace zeitApp
{
    /* Da du wahrscheinlich nur eine einzige Instanz des ExcelHandlers in deinem Programm benötigst, 
könnte das Singleton Pattern hier sinnvoll sein. 
Dieses Pattern stellt sicher, dass nur eine Instanz der Klasse existiert und ermöglicht den globalen Zugriff darauf. */
    public class ExcelHandlerSingleton
    {
        private static ExcelHandlerSingleton? _instance;
        private readonly Excel.Application _excelApp;
        private readonly Workbook _workbook;
        private Worksheet _worksheet;
        private int _lastRow;
        private int _lastColumn;
        private Excel.Range _usedRange;
        private Excel.Range _headerRow; 
        private Excel.Range _footer;
        private string[] _headlines =
        {
            "Datum",
            "Start",
            "Ende",
            "Ist",
            "Ist ab 18:00",
            "Pause",
            "Ist - Pause",
            "Ist ab 18:00 - Pause",
            "Nachtzeit",
            "Nachtzeit - Pause"
        };

        // privater Konstruktor verhindert direkte Instanziierung
        private ExcelHandlerSingleton(string filePath)
        {
            _excelApp = new Excel.Application();
            _excelApp.Visible = true;
            _workbook = _excelApp.Workbooks.Open(filePath);
        }

        public static ExcelHandlerSingleton GetInstance(string filePath)
        {
            _instance ??= new ExcelHandlerSingleton(filePath);
            return _instance;
        }

        public List<string> GetWorksheetNames()
        {
            List<string> sheetNames = [];
            foreach (Worksheet sheet in _workbook.Worksheets)
            {
                sheetNames.Add(sheet.Name);
            }
            return sheetNames;
        }

        public void SelectWorkSheet(int sheetNumber)
        {
            _worksheet = _workbook.Worksheets[sheetNumber];
            GetLastRowOfThisWorksheet();
        }

        // Weitere Methoden wie in der vorherigen Implementierung
        public List<WorkDay> LoadWorkDaysFromSheet()
        {
            if (_worksheet == null)
            {
                MessageBox.Show("Es wurde keine Tabelle ausgewählt");
                return [];
            }
            var workingDays = new List<WorkDay>();
            Excel.Range dataReadRng = _worksheet.Range[_worksheet.Cells[2, 1], _worksheet.Cells[this._lastRow, 3]];
            // Beispielsweise von Zeile 2 bis n die Arbeitszeiten laden
            foreach(Excel.Range row in dataReadRng.Rows)
            {
                DateTime date = DateTime.Now;
                Double? startTimeAsDouble = null;
                Double? endTimeAsDouble = null;
                WorkDay workDay;
                int col = 1;
                foreach(Excel.Range cell in row.Cells)
                {
                    switch (col)
                    {
                        case 1:
                            date = (DateTime) cell.Value;
                            break;
                        case 2:
                            startTimeAsDouble = cell.Value2 ?? cell.Value2;
                            break;
                        case 3:
                            endTimeAsDouble = cell.Value2 ?? cell.Value2;
                            break;
                    }

                    col++;
                }
                if(startTimeAsDouble == null || endTimeAsDouble == null)
                {
                    workDay = new WorkDay(date);
                }
                else
                {
                    DateTime startTime = DateTime.FromOADate(date.ToOADate() + (double)startTimeAsDouble);
                    DateTime endTime = DateTime.FromOADate(date.ToOADate() + (double)endTimeAsDouble);
                    endTime = endTime < startTime ? endTime.AddDays(1) : endTime;
                    workDay = new WorkDay(date, startTime, endTime); ;
                }
                workingDays.Add(workDay);
            }
            return workingDays;
           
        }

        public void WriteWorkDaysToSheet(WorkMonth workMonth)
        {
            WriteHeadLine();
            WriteDataToUsedRange(workMonth);
            WriteFooterRow(workMonth);

        }

        private void WriteDataToUsedRange(WorkMonth workMonth)
        {
            GetUsedRangeOfThisWorksheet();

            // Definiere Farben für das Zebramuster
            var evenRowColor = Excel.XlRgbColor.rgbLightGray;
            var oddRowColor = Excel.XlRgbColor.rgbWhite;
            int listItemIndex = 0;
            foreach (Excel.Range usedRngRow in this._usedRange.Rows)
            {
                
                var workDay = workMonth.WorkDays[listItemIndex];
                if (workDay.IsWorkingDay)
                {
                    usedRngRow.Cells[4].Value = workDay.GetTotalWorkTimeFormatted();
                    usedRngRow.Cells[5].Value = workDay.GetTotalWorkTimeFrom6Formatted();
                    usedRngRow.Cells[6].Value = workDay.GetBreakTimeFormatted();
                    usedRngRow.Cells[7].Value = workDay.GetWorkTimeWithBreakFormatted();
                    usedRngRow.Cells[8].Value = workDay.GetWorkTimeFrom6WithBreakFormatted();
                    usedRngRow.Cells[9].Value = workDay.GetNightWorkTimeFormatted();
                    usedRngRow.Cells[10].Value = workDay.GetNightWorkTimeWithBreakFormatted();                  
                }

                if (listItemIndex % 2 == 0)
                {
                    usedRngRow.Interior.Color = evenRowColor;
                }
                else
                {
                    usedRngRow.Interior.Color = oddRowColor;
                }
                listItemIndex++;
            }
            Excel.Borders borders = _usedRange.Borders;
            borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;

            // Spaltenbreite nur für den benutzten Bereich anpassen
            GetLastColumnOfThisWorksheet();
            Excel.Range wholeSheet = _worksheet.Range[_worksheet.Cells[1, 1],_worksheet.Cells[_lastRow, _lastColumn]];
            wholeSheet.Columns.AutoFit();
        }

        private void WriteFooterRow(WorkMonth workMonth)
        {
            GetFooterRowOfThisWorksheet();

            // Ändere die Schriftgröße
            _footer.Font.Size = 11;
            // Setze die Hintergrundfarbe der Überschriftenzeile
            _footer.Interior.Color = Excel.XlRgbColor.rgbLightGreen;
            _footer.Cells[1].Value2 = "Summen: ";
            _footer.Cells[4].Value = workMonth.CalculateTotalWorkTime();
            _footer.Cells[5].Value = workMonth.CalculateTotalWorkTimeFrom6();
            _footer.Cells[6].Value = workMonth.CalculateTotalBreakTime();
            _footer.Cells[7].Value = workMonth.CalculateWorkTimeWithBreak();
            _footer.Cells[8].Value = workMonth.CalculateWorkTimeFrom6WithBreak();
            _footer.Cells[9].Value = workMonth.CalculateNightWorkTime();
            _footer.Cells[10].Value = workMonth.CalculateNightWorkTimeWithBreak();
        }

        private void WriteHeadLine()
        {
            GetHeaderRowOfThisWorksheet();
            int col = 0;
            foreach(Excel.Range headerCell  in this._headerRow.Cells)
            {
                headerCell.Value2 = _headlines[col];
                col++;
            }

            // Setze die Überschriften fett
            _headerRow.Font.Bold = true;

            // Ändere die Schriftgröße
            _headerRow.Font.Size = 12;

            // Setze die Hintergrundfarbe der Überschriftenzeile
            _headerRow.Interior.Color = Excel.XlRgbColor.rgbLightBlue;

            // Rahmen um die Überschriftenzeile setzen
            _headerRow.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            _headerRow.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            _headerRow.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            _headerRow.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            _headerRow.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
        }

        public void SaveAndClose()
        {
            _workbook.Save();
            _workbook.Close();
            _excelApp.Quit();


            // Ressourcen freigeben
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_worksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelApp);
        }

        public void FormatWorksheet()
        {
                        // ************** Header-Row ****************
            // Hole die erste Zeile (Überschriftenzeile)
            Excel.Range headerRow = _worksheet.Range[_worksheet.Cells[1, 1], _worksheet.Cells[1, _lastColumn]];

            

            // *********** Table - UsedRange ***********

            // Definiere Farben für das Zebramuster
            var evenRowColor = Excel.XlRgbColor.rgbLightGray;
            var oddRowColor = Excel.XlRgbColor.rgbWhite;

            // Schleife durch die Zeilen und wende die Farben nur auf die benutzten Spalten an
            for (int i = 2; i <= _lastRow; i++) // Beginne bei 2, da die erste Zeile die Überschrift ist
            {
                Excel.Range row = _worksheet.Range[_worksheet.Cells[i, 1], _worksheet.Cells[i, _lastColumn]]; // Nur benutzte Spalten

                if (i % 2 == 0)
                {
                    row.Interior.Color = evenRowColor; // Hintergrundfarbe für gerade Zeilen
                }
                else
                {
                    row.Interior.Color = oddRowColor;  // Hintergrundfarbe für ungerade Zeilen
                }
            }

            

            // *********** Borders UsedRange ************
            // Rahmenlinien um den benutzten Bereich setzen
            
            Excel.Borders borders = _usedRange.Borders;
            borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;

            // Spaltenbreite nur für den benutzten Bereich anpassen
            //usedRange.Columns.AutoFit();
        }

        private void GetLastRowOfThisWorksheet()
        {                 
            this._lastRow = _worksheet.Cells[_worksheet.Rows.Count, 1].End(Excel.XlDirection.xlUp).Row; // Letzte Zeile mit Date
        }

        private void GetLastColumnOfThisWorksheet()
        {
            this._lastColumn = _worksheet.Cells[1, _worksheet.Columns.Count].End(Excel.XlDirection.xlToLeft).Column; // Letzte Spalte mit Daten
        }

        private void GetHeaderRowOfThisWorksheet()
        {
            this._headerRow = _worksheet.Range[_worksheet.Cells[1, 1], _worksheet.Cells[1, this._headlines.Length]];
        }

        private void GetFooterRowOfThisWorksheet()
        {
            this._footer = _worksheet.Range[_worksheet.Cells[_lastRow + 1, 1], _worksheet.Cells[_lastRow + 1, this._headlines.Length]];
        }

        private void GetUsedRangeOfThisWorksheet()
        {
            this._usedRange = _worksheet.Range[_worksheet.Cells[2, 1], _worksheet.Cells[_lastRow, this._headlines.Length]];
        }

        public void ExportWorksheetAsPDF(Excel.Worksheet ws)
        {
            string pdfPath = @"F:\D+P_Naumann\pdfs\";
            ws.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfPath + ws.Name + ".pdf");
        }
    }
}
