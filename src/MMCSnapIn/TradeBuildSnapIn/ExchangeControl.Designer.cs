namespace com.tradewright.tradebuildsnapin {
    partial class ExchangeControl {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.NameText = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.TimezoneCombo = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.NotesText = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.TimezoneText = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // NameText
            // 
            this.NameText.Location = new System.Drawing.Point(75, 7);
            this.NameText.Name = "NameText";
            this.NameText.Size = new System.Drawing.Size(182, 20);
            this.NameText.TabIndex = 0;
            this.NameText.TextChanged += new System.EventHandler(this.NameText_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Name";
            // 
            // TimezoneCombo
            // 
            this.TimezoneCombo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.TimezoneCombo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.TimezoneCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.TimezoneCombo.FormattingEnabled = true;
            this.TimezoneCombo.Location = new System.Drawing.Point(75, 33);
            this.TimezoneCombo.Name = "TimezoneCombo";
            this.TimezoneCombo.Size = new System.Drawing.Size(182, 21);
            this.TimezoneCombo.TabIndex = 2;
            this.TimezoneCombo.SelectedIndexChanged += new System.EventHandler(this.TimezoneCombo_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 36);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Timezone";
            // 
            // NotesText
            // 
            this.NotesText.Location = new System.Drawing.Point(75, 86);
            this.NotesText.Multiline = true;
            this.NotesText.Name = "NotesText";
            this.NotesText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.NotesText.Size = new System.Drawing.Size(328, 102);
            this.NotesText.TabIndex = 3;
            this.NotesText.TextChanged += new System.EventHandler(this.NotesText_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 89);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Notes";
            // 
            // TimezoneText
            // 
            this.TimezoneText.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.TimezoneText.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.TimezoneText.Location = new System.Drawing.Point(75, 60);
            this.TimezoneText.Name = "TimezoneText";
            this.TimezoneText.Size = new System.Drawing.Size(328, 20);
            this.TimezoneText.TabIndex = 6;
            this.TimezoneText.TabStop = false;
            // 
            // ExchangeControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.TimezoneText);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.NotesText);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.TimezoneCombo);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.NameText);
            this.Name = "ExchangeControl";
            this.Size = new System.Drawing.Size(413, 200);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox NameText;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox TimezoneCombo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox NotesText;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox TimezoneText;
    }
}
