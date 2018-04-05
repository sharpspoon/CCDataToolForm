using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.IO.Compression;

namespace DataAnalysisTool
{
    public partial class CheckTools : Form
    {
        public DataAnalysisTool ccdatatoolform;
        DataAnalysisTool ccd = new DataAnalysisTool();
        
        public CheckTools()
        {
            InitializeComponent();
        }
        
        private void checkButton1_Click(object sender, EventArgs e)
        {
            if (ctTextBox2.Text.Length == 0)
            {
                MessageBox.Show("You did not enter a column name!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                //return;
            }
            else
            for (int i = 0; i < ccd.importedfileDataGridView.Rows.Count; i++)
            {
                MessageBox.Show("you got in the loop", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                var value = ccd.importedfileDataGridView.Rows[i].Cells[ctTextBox2.Text].Value.ToString();
                if ((value.Length != 8) && (value != null) && (value != ""))
                {
                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "Make sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    return;
                }
            }

            for (int i = 0; i < ccd.importedfileDataGridView.Rows.Count; i++)
            {
                try
                {
                    var value2 = ccd.importedfileDataGridView.Rows[i].Cells[ctTextBox2.Text].Value.ToString();
                    int year = int.Parse(value2.Substring(0, 4));
                    int month = int.Parse(value2.Substring(4, 2));
                    int day = int.Parse(value2.Substring(6, 2));

                    if (year > 2200)
                    {
                        MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Dates are OK");
                    return;
                }
            }
        }

        private void ctTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void CheckTools_Load(object sender, EventArgs e)
        {

        }
    }
}
