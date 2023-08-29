namespace com.tradewright.tradebuildsnapin {
    partial class DatabaseControl {
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
            this.TypeCombo = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.ServerText = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.DatabaseText = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.UsernameText = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.PasswordText = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // TypeCombo
            // 
            this.TypeCombo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.TypeCombo.FormattingEnabled = true;
            this.TypeCombo.Location = new System.Drawing.Point(69, 22);
            this.TypeCombo.Name = "TypeCombo";
            this.TypeCombo.Size = new System.Drawing.Size(156, 21);
            this.TypeCombo.TabIndex = 0;
            this.TypeCombo.SelectedIndexChanged += new System.EventHandler(this.TypeCombo_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(31, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Type";
            // 
            // ServerText
            // 
            this.ServerText.Location = new System.Drawing.Point(69, 49);
            this.ServerText.Name = "ServerText";
            this.ServerText.Size = new System.Drawing.Size(156, 20);
            this.ServerText.TabIndex = 2;
            this.ServerText.TextChanged += new System.EventHandler(this.ServerText_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 52);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(38, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Server";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 78);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Database";
            // 
            // DatabaseText
            // 
            this.DatabaseText.Location = new System.Drawing.Point(69, 75);
            this.DatabaseText.Name = "DatabaseText";
            this.DatabaseText.Size = new System.Drawing.Size(156, 20);
            this.DatabaseText.TabIndex = 4;
            this.DatabaseText.TextChanged += new System.EventHandler(this.DatabaseText_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 104);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Username";
            // 
            // UsernameText
            // 
            this.UsernameText.Location = new System.Drawing.Point(69, 101);
            this.UsernameText.Name = "UsernameText";
            this.UsernameText.Size = new System.Drawing.Size(156, 20);
            this.UsernameText.TabIndex = 6;
            this.UsernameText.TextChanged += new System.EventHandler(this.UsernameText_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 130);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 13);
            this.label5.TabIndex = 9;
            this.label5.Text = "Password";
            // 
            // PasswordText
            // 
            this.PasswordText.Location = new System.Drawing.Point(69, 127);
            this.PasswordText.Name = "PasswordText";
            this.PasswordText.PasswordChar = '*';
            this.PasswordText.Size = new System.Drawing.Size(156, 20);
            this.PasswordText.TabIndex = 8;
            this.PasswordText.TextChanged += new System.EventHandler(this.PasswordText_TextChanged);
            // 
            // DatabaseControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.label5);
            this.Controls.Add(this.PasswordText);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.UsernameText);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.DatabaseText);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.ServerText);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TypeCombo);
            this.Name = "DatabaseControl";
            this.Size = new System.Drawing.Size(251, 169);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox TypeCombo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox ServerText;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox DatabaseText;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox UsernameText;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox PasswordText;

    }
}
