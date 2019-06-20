namespace SAPDataAnalysisTool
{
    partial class Importformat
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
            this.components = new System.ComponentModel.Container();
            this.ifSelect2 = new System.Windows.Forms.ComboBox();
            this.cCDataToolBindingSource = new System.Windows.Forms.BindingSource(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.cCDataToolBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // ifSelect2
            // 
            this.ifSelect2.Cursor = System.Windows.Forms.Cursors.Default;
            this.ifSelect2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ifSelect2.FormattingEnabled = true;
            this.ifSelect2.Location = new System.Drawing.Point(399, 121);
            this.ifSelect2.Name = "ifSelect2";
            this.ifSelect2.Size = new System.Drawing.Size(111, 21);
            this.ifSelect2.TabIndex = 33;
            this.ifSelect2.SelectedIndexChanged += new System.EventHandler(this.ifSelect_SelectedIndexChanged);
            // 
            // cCDataToolBindingSource
            // 
            this.cCDataToolBindingSource.DataSource = typeof(SAPDataAnalysisTool);
            // 
            // Importformat
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(909, 262);
            this.Controls.Add(this.ifSelect2);
            this.Name = "Importformat";
            this.Text = "Importformat";
            this.Load += new System.EventHandler(this.Importformat_Load);
            ((System.ComponentModel.ISupportInitialize)(this.cCDataToolBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.ComboBox ifSelect2;
        private System.Windows.Forms.BindingSource cCDataToolBindingSource;
    }
}