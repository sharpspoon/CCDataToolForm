namespace CCDataImportTool
{
    partial class CheckTools
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CheckTools));
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.ctTextBox2 = new System.Windows.Forms.TextBox();
            this.checkButton1 = new System.Windows.Forms.Button();
            this.checkToolsLabel = new System.Windows.Forms.Label();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.Color.AliceBlue;
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.ctTextBox2);
            this.groupBox2.Controls.Add(this.checkButton1);
            this.groupBox2.Location = new System.Drawing.Point(73, 100);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(144, 62);
            this.groupBox2.TabIndex = 24;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Date Converter:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 19);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(74, 13);
            this.label4.TabIndex = 25;
            this.label4.Text = "Column name:";
            // 
            // ctTextBox2
            // 
            this.ctTextBox2.Location = new System.Drawing.Point(6, 35);
            this.ctTextBox2.Name = "ctTextBox2";
            this.ctTextBox2.Size = new System.Drawing.Size(74, 20);
            this.ctTextBox2.TabIndex = 20;
            this.ctTextBox2.TextChanged += new System.EventHandler(this.ctTextBox2_TextChanged);
            // 
            // checkButton1
            // 
            this.checkButton1.Cursor = System.Windows.Forms.Cursors.Default;
            this.checkButton1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkButton1.Image = global::CCDataImportTool.Properties.Resources.StatusAnnotations_Play_32xLG_color;
            this.checkButton1.Location = new System.Drawing.Point(86, 16);
            this.checkButton1.Name = "checkButton1";
            this.checkButton1.Size = new System.Drawing.Size(48, 40);
            this.checkButton1.TabIndex = 21;
            this.checkButton1.UseVisualStyleBackColor = true;
            this.checkButton1.Click += new System.EventHandler(this.checkButton1_Click);
            // 
            // checkToolsLabel
            // 
            this.checkToolsLabel.AutoSize = true;
            this.checkToolsLabel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.checkToolsLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkToolsLabel.Location = new System.Drawing.Point(47, 9);
            this.checkToolsLabel.Name = "checkToolsLabel";
            this.checkToolsLabel.Size = new System.Drawing.Size(195, 37);
            this.checkToolsLabel.TabIndex = 25;
            this.checkToolsLabel.Text = "Check Tools";
            // 
            // CheckTools
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.AliceBlue;
            this.ClientSize = new System.Drawing.Size(287, 194);
            this.Controls.Add(this.checkToolsLabel);
            this.Controls.Add(this.groupBox2);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "CheckTools";
            this.Text = "Check Tools";
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label checkToolsLabel;
        public System.Windows.Forms.TextBox ctTextBox2;
        public System.Windows.Forms.Button checkButton1;
    }
}