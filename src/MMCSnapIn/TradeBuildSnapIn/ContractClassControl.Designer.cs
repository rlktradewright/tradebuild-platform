namespace com.tradewright.tradebuildsnapin {
    partial class ContractClassControl {
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
            this.components = new System.ComponentModel.Container();
            this.label3 = new System.Windows.Forms.Label();
            this.NotesText = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SecTypeCombo = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.NameText = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.SwitchDayText = new System.Windows.Forms.TextBox();
            this.TickSizeText = new System.Windows.Forms.TextBox();
            this.TickValueText = new System.Windows.Forms.TextBox();
            this.SessionStartText = new System.Windows.Forms.TextBox();
            this.SessionEndText = new System.Windows.Forms.TextBox();
            this.CurrencyCombo = new System.Windows.Forms.ComboBox();
            this.CurrencyText = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 198);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(35, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "Notes";
            // 
            // NotesText
            // 
            this.NotesText.Location = new System.Drawing.Point(77, 195);
            this.NotesText.Multiline = true;
            this.NotesText.Name = "NotesText";
            this.NotesText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.NotesText.Size = new System.Drawing.Size(328, 102);
            this.NotesText.TabIndex = 8;
            this.NotesText.TextChanged += new System.EventHandler(this.NotesText_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(49, 13);
            this.label2.TabIndex = 9;
            this.label2.Text = "Sec type";
            // 
            // SecTypeCombo
            // 
            this.SecTypeCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.SecTypeCombo.FormattingEnabled = true;
            this.SecTypeCombo.Location = new System.Drawing.Point(75, 35);
            this.SecTypeCombo.Name = "SecTypeCombo";
            this.SecTypeCombo.Size = new System.Drawing.Size(182, 21);
            this.SecTypeCombo.TabIndex = 1;
            this.SecTypeCombo.SelectedIndexChanged += new System.EventHandler(this.SecTypeCombo_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 7;
            this.label1.Text = "Name";
            // 
            // NameText
            // 
            this.NameText.Location = new System.Drawing.Point(75, 9);
            this.NameText.Name = "NameText";
            this.NameText.Size = new System.Drawing.Size(182, 20);
            this.NameText.TabIndex = 0;
            this.NameText.TextChanged += new System.EventHandler(this.NameText_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(9, 65);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(49, 13);
            this.label4.TabIndex = 13;
            this.label4.Text = "Currency";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(9, 91);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(49, 13);
            this.label5.TabIndex = 15;
            this.label5.Text = "Tick size";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(9, 117);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(57, 13);
            this.label6.TabIndex = 17;
            this.label6.Text = "Tick value";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(9, 143);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(59, 13);
            this.label7.TabIndex = 19;
            this.label7.Text = "Switch day";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(141, 143);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(264, 13);
            this.label8.TabIndex = 20;
            this.label8.Text = "(Days before contract expiry to switch to next contract)";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(109, 172);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(29, 13);
            this.label9.TabIndex = 22;
            this.label9.Text = "Start";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(191, 172);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(26, 13);
            this.label10.TabIndex = 24;
            this.label10.Text = "End";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(9, 172);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(74, 13);
            this.label11.TabIndex = 25;
            this.label11.Text = "Session times:";
            // 
            // toolTip1
            // 
            this.toolTip1.AutomaticDelay = 0;
            this.toolTip1.AutoPopDelay = 5000;
            this.toolTip1.InitialDelay = 500;
            this.toolTip1.IsBalloon = true;
            this.toolTip1.ReshowDelay = 100;
            this.toolTip1.ShowAlways = true;
            this.toolTip1.ToolTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            // 
            // SwitchDayText
            // 
            this.SwitchDayText.Location = new System.Drawing.Point(74, 140);
            this.SwitchDayText.Name = "SwitchDayText";
            this.SwitchDayText.Size = new System.Drawing.Size(65, 20);
            this.SwitchDayText.TabIndex = 5;
            this.SwitchDayText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.toolTip1.SetToolTip(this.SwitchDayText, "The number of days before contract expiry to switch to the next contract");
            this.SwitchDayText.TextChanged += new System.EventHandler(this.SwitchDayText_TextChanged);
            // 
            // TickSizeText
            // 
            this.TickSizeText.Location = new System.Drawing.Point(74, 88);
            this.TickSizeText.Name = "TickSizeText";
            this.TickSizeText.Size = new System.Drawing.Size(65, 20);
            this.TickSizeText.TabIndex = 3;
            this.TickSizeText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.TickSizeText.TextChanged += new System.EventHandler(this.TickSizeText_TextChanged);
            // 
            // TickValueText
            // 
            this.TickValueText.Location = new System.Drawing.Point(74, 114);
            this.TickValueText.Name = "TickValueText";
            this.TickValueText.Size = new System.Drawing.Size(65, 20);
            this.TickValueText.TabIndex = 4;
            this.TickValueText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.TickValueText.TextChanged += new System.EventHandler(this.TickValueText_TextChanged);
            // 
            // SessionStartText
            // 
            this.SessionStartText.Location = new System.Drawing.Point(144, 169);
            this.SessionStartText.Name = "SessionStartText";
            this.SessionStartText.Size = new System.Drawing.Size(39, 20);
            this.SessionStartText.TabIndex = 6;
            this.SessionStartText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.SessionStartText.TextChanged += new System.EventHandler(this.SessionStartText_TextChanged);
            // 
            // SessionEndText
            // 
            this.SessionEndText.Location = new System.Drawing.Point(223, 169);
            this.SessionEndText.Name = "SessionEndText";
            this.SessionEndText.Size = new System.Drawing.Size(39, 20);
            this.SessionEndText.TabIndex = 7;
            this.SessionEndText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.SessionEndText.TextChanged += new System.EventHandler(this.SessionEndText_TextChanged);
            // 
            // CurrencyCombo
            // 
            this.CurrencyCombo.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.CurrencyCombo.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.CurrencyCombo.FormattingEnabled = true;
            this.CurrencyCombo.Location = new System.Drawing.Point(75, 62);
            this.CurrencyCombo.Name = "CurrencyCombo";
            this.CurrencyCombo.Size = new System.Drawing.Size(65, 21);
            this.CurrencyCombo.TabIndex = 2;
            this.CurrencyCombo.SelectedIndexChanged += new System.EventHandler(this.CurrencyCombo_SelectedIndexChanged);
            // 
            // CurrencyText
            // 
            this.CurrencyText.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.CurrencyText.Location = new System.Drawing.Point(152, 63);
            this.CurrencyText.Name = "CurrencyText";
            this.CurrencyText.Size = new System.Drawing.Size(252, 20);
            this.CurrencyText.TabIndex = 26;
            this.CurrencyText.TabStop = false;
            // 
            // ContractClassControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.CurrencyText);
            this.Controls.Add(this.CurrencyCombo);
            this.Controls.Add(this.SessionEndText);
            this.Controls.Add(this.SessionStartText);
            this.Controls.Add(this.SwitchDayText);
            this.Controls.Add(this.TickValueText);
            this.Controls.Add(this.TickSizeText);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.NotesText);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.SecTypeCombo);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.NameText);
            this.Name = "ContractClassControl";
            this.Size = new System.Drawing.Size(412, 306);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox NotesText;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox SecTypeCombo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox NameText;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.TextBox TickSizeText;
        private System.Windows.Forms.TextBox TickValueText;
        private System.Windows.Forms.TextBox SwitchDayText;
        private System.Windows.Forms.TextBox SessionStartText;
        private System.Windows.Forms.TextBox SessionEndText;
        private System.Windows.Forms.ComboBox CurrencyCombo;
        private System.Windows.Forms.TextBox CurrencyText;
    }
}
