using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace zeitApp
{
    public class WorksheetSelector
    {
        public int SelectExcelWorksheet(IEnumerable<string> sheetNames)
        {
            
            int selectedWorksheetNumber = 0;
            using (var inputForm = new SheetInputForm())
            {
                inputForm.Text = "Wähle das entsprechende Tabellenblatt aus!";
                inputForm.LoadSheetNames(sheetNames);


                if (inputForm.ShowDialog() == DialogResult.OK)
                {
                    selectedWorksheetNumber = inputForm.SelectedSheetNumber;
                }

                if (selectedWorksheetNumber == 0)
                {
                    MessageBox.Show("Das angegebene Tabellenblatt existiert nicht.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return 0;
                }
            }

            // Beginn von Worksheets = 1; ListBoxItems = 0
            return selectedWorksheetNumber + 1;
        }

    }
}
