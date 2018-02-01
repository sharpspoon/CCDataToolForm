using System;
using System.Globalization;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.DirectoryServices.ActiveDirectory;
using System.Web;

namespace CCDataImportTool
{
    public partial class Form1 : Form
    {
        //------------------OPEN/SAVE CSV START------------------------------------------------------
        private void menu_Open_Csv_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "CSV | *.csv", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        dataGridView1.DataSource = ReadCsv(ofd.FileName);
                    textBox1.Text = ofd.FileName;
                    textBox7.Text = dataGridView1.Rows.Count.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public DataTable ReadCsv(string fileName)
        {
            DataTable dt = new DataTable("Data");
            using (OleDbConnection cn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"" +
                Path.GetDirectoryName(fileName) + "\";Extended Properties='text;HDR=yes;FMT=Delimited(,)';"))
            {
                using (OleDbCommand cmd = new OleDbCommand(string.Format("select * from [{0}]", new FileInfo(fileName).Name), cn))
                {
                    cn.Open();
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
                    {
                        adapter.Fill(dt);
                    }
                }
            }
            return dt;
        }
        protected void menu_Save_Csv_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "CSV|*.csv";
            if (this.saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // create one file gridview.csv in writing mode using streamwriter
                StreamWriter sw = new StreamWriter("c:\\gridview.csv");
                // now add the gridview header in csv file suffix with "," delimeter except last one
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    sw.Write(dataGridView1.Columns[i].HeaderText);
                    if (i != dataGridView1.Columns.Count)
                    {
                        sw.Write(",");
                    }
                }
                // add new line
                sw.Write(sw.NewLine);
                // iterate through all the rows within the gridview
                foreach (DataGridViewRow dr in dataGridView1.Rows)
                {
                    // iterate through all colums of specific row
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        // write particular cell to csv file
                        sw.Write(dr.Cells[i]);
                        if (i != dataGridView1.Columns.Count)
                        {
                            sw.Write(",");
                        }
                    }
                    // write new line
                    sw.Write(sw.NewLine);
                }
                // flush from the buffers.
                sw.Flush();
                // closes the file
                sw.Close();
            }
        }

        //------------------OPEN/SAVE CSV END------------------------------------------------------

        //------------------OPEN/SAVE XML START------------------------------------------------------

        private void menu_Open_Xml_Click(object sender, EventArgs e)
        {
            try
            {
                DataSet dataSet = new DataSet();
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "XML | *.xml", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        dataSet.ReadXml(ofd.FileName);
                    dataGridView1.DataSource = dataSet.Tables[0];

                    textBox1.Text = ofd.FileName;
                    textBox7.Text = dataGridView1[0,dataGridView1.Rows.Count-1].Value.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void menu_Save_Xml_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "XML|*.xml";
            if (this.saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DataTable dt = (DataTable)this.dataGridView1.DataSource;
                dt.WriteXml(this.saveFileDialog1.FileName, XmlWriteMode.WriteSchema);
            }
        }

        //------------------OPEN/SAVE XML END------------------------------------------------------

        //------------------OPEN/SAVE XLS START------------------------------------------------------

        private void menu_Open_Xls_Click(object sender, EventArgs e)
        {

        }
        //------------------OPEN/SAVE XLS END------------------------------------------------------

        //------------------PRINT DOCUMENT START------------------------------------------------------

        Bitmap bitmap;
        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0 || dataGridView1.Rows == null)
            {
                MessageBox.Show("No data to print", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                //Resize DataGridView to full height.
                int height = dataGridView1.Height;
                dataGridView1.Height = dataGridView1.RowCount * dataGridView1.RowTemplate.Height;

                //Create a Bitmap and draw the DataGridView on it.
                bitmap = new Bitmap(this.dataGridView1.Width, this.dataGridView1.Height);
                dataGridView1.DrawToBitmap(bitmap, new Rectangle(0, 0, this.dataGridView1.Width, this.dataGridView1.Height));

                //Resize DataGridView back to original height.
                dataGridView1.Height = height;

                //Show the Print Preview Dialog.
                printPreviewDialog1.Document = printDocument1;
                printPreviewDialog1.PrintPreviewControl.Zoom = 1;
                printPreviewDialog1.ShowDialog();
            }
        }
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //Print the contents.
            e.Graphics.DrawImage(bitmap, 0, 0);
        }

        //------------------PRINT DOCUMENT END------------------------------------------------------

        //------------------EXIT APP ACTION START------------------------------------------------------

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult result = MessageBox.Show("Do you really want to exit?", "CCDataTool", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    Environment.Exit(0);
                }
                else
                {
                    e.Cancel = true;
                }
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void exitToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        //------------------EXIT APP ACTION END------------------------------------------------------

        //------------------ENVIRONMENT MENU START------------------------------------------------------

        private void env_Click1(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://hmigexttest2.callidusinsurance.net/ICM");
        }

        private void env_Click2(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://hmigexttest3.callidusinsurance.net/ICM");
        }

        //------------------ENVIRONMENT MENU END------------------------------------------------------

        //------------------DATE CONVERTER START------------------------------------------------------

        private void dateConvert_Click1(object sender, EventArgs e)
        {
            try
            {
                string newPattern = "yyyyMMdd";
                DateTime thisDate1 = new DateTime();
                dataGridView1.Columns[textBox2.Text].DefaultCellStyle.Format = thisDate1.ToString(newPattern);
            }
            catch (Exception ex)
            {
                if (textBox2.Text.Length == 0)
                {
                    MessageBox.Show("You did not enter a column name!\r\nThe operation will now cancel.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    return;
                }
                MessageBox.Show(ex.Message, "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
        }

        //------------------DATE CONVERTER END------------------------------------------------------

        //------------------NULL CHECKER START------------------------------------------------------

        private void nullChecker_Click(object sender, EventArgs e)
        {
            if (textBox6.Text.Length == 0)
            {
                MessageBox.Show("You did not enter a column name!\r\nThe operation will now cancel.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                try
                {
                    var value = dataGridView1.Rows[i].Cells[textBox6.Text].Value.ToString();
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        MessageBox.Show("NULL value found in column " + "'" + textBox6.Text + "'" + " at line " + dataGridView1.Rows[i + 1]);
                        return;
                    }
                }
                catch (Exception)
                {
                    // If we have reached this far, then none of the cells were empty.
                    MessageBox.Show("No NULL values found in column " + "'" + textBox6.Text + "'");
                    return;
                }
            }
        }
        private void textBox6_TextChanged(object sender, EventArgs e)
        {
        }

        //------------------NULL CHECKER END------------------------------------------------------

        //------------------SPECIAL CHARACTER CHECKER START------------------------------------------------------

        private void button3_Click(object sender, EventArgs e)
        {
            String searchValue = comboBox1.Text;
            string specialBoxFill = textBox5.Text;
            if (textBox5.Text.Length == 0)
            {
                MessageBox.Show("You did not enter a column name!\r\nThe operation will now cancel.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            if (comboBox1.Text.Length == 0)
            {
                MessageBox.Show("You did not select a special character!\r\nThe operation will now cancel.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    if (row.Cells[textBox5.Text].Value.ToString().Contains(comboBox1.Text))
                    {
                        MessageBox.Show("'" + comboBox1.Text + "'" + " WAS found in the column " + "'" + textBox5.Text + "'", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                        return;
                    }
                }
                catch
                {
                    MessageBox.Show("'"+comboBox1.Text+"'" + " WAS NOT  found in column "+"'"+textBox5.Text+"'", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return;
                }



                
                }

            
        }

        //------------------SPECIAL CHARACTER CHECKER END------------------------------------------------------

        //------------------CELL LENGTH CHECKER START------------------------------------------------------

        private void button4_Click(object sender, EventArgs e)
        {
            {
                try
                {
                    DataGridViewColumn column = dataGridView1.Columns[textBox3.Text];
                    MessageBox.Show(column.Name + " must be " + textBox4.Text + " Digit(s) Long!");
                }
                catch (Exception ex)
                {
                    if (textBox3.Text.Length == 0)
                    {
                        MessageBox.Show("You did not enter a column name!\r\nThe operation will now cancel.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }
                    if (textBox4.Text.Length == 0)
                    {
                        MessageBox.Show("You did not enter a length!\r\nThe operation will now cancel.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }
                    MessageBox.Show(ex.Message, "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        //------------------CELL LENGTH CHECKER END------------------------------------------------------

        //------------------ABOUT START------------------------------------------------------

        private void menu_About_Click(object sender, EventArgs e)
        {
            About about = new About();
            about.Show();
        }

        //------------------ABOUT END------------------------------------------------------

        //------------------CC LOGO CLICK START------------------------------------------------------
        private void ccLogo_Click1(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://calliduscloud.com");
        }


        //------------------CC LOGO CLICK END------------------------------------------------------

        //------------------MEDICARE CHECKER START------------------------------------------------------

        private void medicareButton_Click(object sender, EventArgs e)
        {
            String searchValue = comboBox1.Text;
            string specialBoxFill = textBox5.Text;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                try
                {
                    if (dataGridView1.ColumnCount != 37)
                    {
                        MessageBox.Show("Medicare files need 37 columns. You have " + dataGridView1.ColumnCount + ".", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }
                    var value0 = dataGridView1.Rows[i].Cells[0].Value.ToString();
                    //var value1 = dataGridView1.Rows[i].Cells[1].Value.ToString();
                    //var value2 = dataGridView1.Rows[i].Cells[2].Value.ToString();
                    //var value3 = dataGridView1.Rows[i].Cells[3].Value.ToString();
                    //var value6 = dataGridView1.Rows[i].Cells[6].Value.ToString();
                    //var value8 = dataGridView1.Rows[i].Cells[8].Value.ToString();
                    //var value10 = dataGridView1.Rows[i].Cells[10].Value.ToString();
                    //var value23 = dataGridView1.Rows[i].Cells[23].Value.ToString();
                    if (string.IsNullOrWhiteSpace(value0))
                    {
                        MessageBox.Show("NULL value found in column #1 (CustomerId)  at line " + (i + 1) + " This is a required field.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }

                    //if (string.IsNullOrWhiteSpace(value1))
                    //{
                    //    MessageBox.Show("NULL value found in column #2 (ContractNbr)  at line " + dataGridView1.Rows[i + 1] + " This is a required field.");
                    //    return;
                    //}

                    //if (string.IsNullOrWhiteSpace(value2))
                    //{
                    //    MessageBox.Show("NULL value found in column #3 (PBP)  at line " + dataGridView1.Rows[i + 1] + " This is a required field.");
                    //    return;
                    //}

                    //if (string.IsNullOrWhiteSpace(value3))
                    //{
                    //    MessageBox.Show("NULL value found in column #4 (HICN)  at line " + dataGridView1.Rows[i + 1] + " This is a required field.");
                    //    return;
                    //}

                    //if (string.IsNullOrWhiteSpace(value6))
                    //{
                    //    MessageBox.Show("NULL value found in column #7 (DatEff)  at line " + dataGridView1.Rows[i + 1] + " This is a required field.");
                    //    return;
                    //}

                    //if (string.IsNullOrWhiteSpace(value8))
                    //{
                    //    MessageBox.Show("NULL value found in column #9 (AppSignedDate)  at line " + dataGridView1.Rows[i + 1] + " This is a required field.");
                    //    return;
                    //}

                    //if (string.IsNullOrWhiteSpace(value10))
                    //{
                    //    MessageBox.Show("NULL value found in column #11 (Holder)  at line " + dataGridView1.Rows[i + 1] + " This is a required field.");
                    //    return;
                    //}

                    //if (string.IsNullOrWhiteSpace(value23))
                    //{
                    //    MessageBox.Show("NULL value found in column #24 (PolState)  at line " + dataGridView1.Rows[i + 1] + " This is a required field.");
                    //    return;
                    //}
                }
                catch (Exception)
                {
                    // If we have reached this far, then none of the cells were empty.
                    MessageBox.Show("Medicare file is OK");
                    return;
                }
            }


        }

        //------------------MEDICARE CHECKER END------------------------------------------------------


        public Form1()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellValidating_Click(object sender, DataGridViewCellValidatingEventArgs e, EventArgs ea)
        {
            string headerText =
                dataGridView1.Columns[e.ColumnIndex].HeaderText;

            // Abort validation if cell is not in the CompanyName column.
            if (!headerText.Equals("CompanyName")) return;

            // Confirm that the cell is not empty.
            if (string.IsNullOrEmpty(e.FormattedValue.ToString()))
            {
                dataGridView1.Rows[e.RowIndex].ErrorText =
                    "Company Name must not be empty";
                e.Cancel = true;
            }
        }
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            ((DataGridViewTextBoxColumn)dataGridView1.Columns["dates"]).MaxInputLength = 6;
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
        }
        private void groupBox1_Enter(object sender, EventArgs e)
        {
        }
        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
        }
        private void groupBox3_Enter(object sender, EventArgs e)
        {
        }
        private void label2_Click_1(object sender, EventArgs e)
        {
        }
        private void testButton_Click(object sender, EventArgs e)
        {
        }
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
        }

        private void xLSToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void form1BindingSource_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, DataGridViewCellValidatingEventArgs e)
        {
            string headerText =
                dataGridView1.Columns[e.ColumnIndex].HeaderText;

            // Abort validation if cell is not in the CompanyName column.
            if (!headerText.Equals("CompanyName")) return;

            // Confirm that the cell is not empty.
            if (string.IsNullOrEmpty(e.FormattedValue.ToString()))
            {
                dataGridView1.Rows[e.RowIndex].ErrorText =
                    "Company Name must not be empty";
                e.Cancel = true;
            }
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            textBox7.Text = dataGridView1.Rows.Count.ToString();
        }


    }
}
