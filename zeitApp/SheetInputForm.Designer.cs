namespace zeitApp
{
    partial class SheetInputForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btn_Ok = new Button();
            lstSheets = new ListBox();
            btn_cancel = new Button();
            SuspendLayout();
            // 
            // btn_Ok
            // 
            btn_Ok.BackColor = Color.OliveDrab;
            btn_Ok.Location = new Point(50, 390);
            btn_Ok.Name = "btn_Ok";
            btn_Ok.Size = new Size(135, 48);
            btn_Ok.TabIndex = 1;
            btn_Ok.Text = "Okay";
            btn_Ok.UseVisualStyleBackColor = false;
            btn_Ok.Click += Button1_Click;
            // 
            // lstSheets
            // 
            lstSheets.FormattingEnabled = true;
            lstSheets.ItemHeight = 30;
            lstSheets.Location = new Point(51, 106);
            lstSheets.Name = "lstSheets";
            lstSheets.Size = new Size(401, 214);
            lstSheets.TabIndex = 2;
            // 
            // btn_cancel
            // 
            btn_cancel.BackColor = Color.Firebrick;
            btn_cancel.Location = new Point(303, 390);
            btn_cancel.Name = "btn_cancel";
            btn_cancel.Size = new Size(135, 48);
            btn_cancel.TabIndex = 3;
            btn_cancel.Text = "Cancel";
            btn_cancel.UseVisualStyleBackColor = false;
            btn_cancel.Click += Button1_Click_1;
            // 
            // SheetInputForm
            // 
            AutoScaleDimensions = new SizeF(12F, 30F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(518, 460);
            Controls.Add(btn_cancel);
            Controls.Add(lstSheets);
            Controls.Add(btn_Ok);
            Name = "SheetInputForm";
            Text = "Tabellenblatt auswählen";
            ResumeLayout(false);
        }

        #endregion
        private Button btn_Ok;
        private ListBox lstSheets;
        private Button btn_cancel;
    }
}