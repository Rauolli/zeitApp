using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace zeitApp
{
    public partial class SheetInputForm : Form
    {
        public int SelectedSheetNumber { get; private set; }
        public SheetInputForm()
        {
            InitializeComponent();
        }

        public void LoadSheetNames(IEnumerable<string> sheetNames)
        {
            lstSheets.Items.Clear();
            foreach (var name in sheetNames)
            {
                lstSheets.Items.Add(name);
            }
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (lstSheets.SelectedItem != null)
            {
                SelectedSheetNumber = lstSheets.SelectedIndex;
                Console.WriteLine(SelectedSheetNumber);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                MessageBox.Show("Bitte wählen Sie ein Blatt aus.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void Button1_Click_1(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}
