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

                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "XML | *.xml", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        dataGridView1.DataSource = ReadXml(ofd.FileName);
                    textBox1.Text = ofd.FileName;
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public DataSet ReadXml(string fileName)
        {
            try
            {
                XmlReader xmlFile;
                xmlFile = XmlReader.Create(Path.GetDirectoryName(fileName), new XmlReaderSettings());
                DataSet ds = new DataSet();
                ds.ReadXml(xmlFile);
                dataGridView1.DataSource = ds.Tables[0];
                return ds;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            DataSet dt = new DataSet("Data");
            return dt;
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



        public Form1()
        {
            InitializeComponent();
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

        private void menu_About_Click(object sender, EventArgs e)
        {
            About about = new About();
            about.Show();
        }

      

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
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


        private void label1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://hmigexttest2.callidusinsurance.net/ICM");

        }

        private void label2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://hmigexttest3.callidusinsurance.net/ICM");

        }

        private void label3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.microsoft.com");

        }

        private void label4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.microsoft.com");

        }

        private void label5_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.microsoft.com");

        }

        private void label6_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.microsoft.com");

        }

        private void label7_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.microsoft.com");

        }

        private void label8_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.microsoft.com");

        }

        private void label9_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.microsoft.com");

        }

        private void cSVToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            var sb = new StringBuilder();

            var headers = dataGridView1.Columns.Cast<DataGridViewColumn>();
            sb.AppendLine(string.Join(",", headers.Select(column => "\"" + column.HeaderText + "\"").ToArray()));

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                var cells = row.Cells.Cast<DataGridViewCell>();
                sb.AppendLine(string.Join(",", cells.Select(cell => "\"" + cell.Value + "\"").ToArray()));
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {

                string newPattern = "yyyyMMdd";
                DateTime thisDate1 = new DateTime();
                dataGridView1.Columns[textBox2.Text].DefaultCellStyle.Format = thisDate1.ToString(newPattern);
            }


                        catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
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
            int counter;
            int WithdrawalTotal=0;
            for (counter = 0; counter < (dataGridView1.Rows.Count);
                counter++)
            {
                if (dataGridView1.Rows[counter].Cells[textBox3.Text].Value
                    != null)
                {
                    if (dataGridView1.Rows[counter].
                        Cells[textBox3.Text].Value.ToString().Length != 0)
                    {
                        WithdrawalTotal += int.Parse(dataGridView1.Rows[counter].
                            Cells[textBox3.Text].Value.ToString());
                        DataGridViewColumn column = dataGridView1.Columns[textBox3.Text];
                        MessageBox.Show(column.Name + " must be " + textBox4.Text + " Digit(s) Long!");
                    }
                }
            }
        }

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
                    MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }


        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellValidating_Click(object sender,
    DataGridViewCellValidatingEventArgs e)
        {
            dataGridView1.Rows[e.RowIndex].ErrorText = "";
            int newInteger;

            // Don't try to validate the 'new row' until finished 
            // editing since there
            // is not any point in validating its initial value.
            if (dataGridView1.Rows[e.RowIndex].IsNewRow) { return; }
            if (!int.TryParse(e.FormattedValue.ToString(),
                out newInteger) || newInteger < 0)
            {
                e.Cancel = true;
                dataGridView1.Rows[e.RowIndex].ErrorText = "the value must be a non-negative integer";
            }
        }

        private void xLSToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
