namespace TradeWright.TradeBuild.Applications.Chart
{
    partial class ChartStylesOrganizer
    {
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.chartStylesOrganizer1 = new Chart.ChartStylesOrganizer();
            this.SuspendLayout();
            // 
            // chartStylesOrganizer1
            // 
            this.chartStylesOrganizer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chartStylesOrganizer1.Location = new System.Drawing.Point(0, 0);
            this.chartStylesOrganizer1.MinimumSize = new System.Drawing.Size(150, 300);
            this.chartStylesOrganizer1.Name = "chartStylesOrganizer1";
            this.chartStylesOrganizer1.Size = new System.Drawing.Size(786, 441);
            this.chartStylesOrganizer1.TabIndex = 0;
            // 
            // ChartStylesOrganizer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(786, 441);
            this.Controls.Add(this.chartStylesOrganizer1);
            this.Name = "ChartStylesOrganizer";
            this.Text = "Manage Chart Styles";
            this.ResumeLayout(false);

        }

        #endregion

        private ChartStylesOrganizer chartStylesOrganizer1;
    }
}