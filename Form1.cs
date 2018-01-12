using System;
using System.Globalization;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;

namespace CCDataImportTool
{
    public partial class Form1 : Form
    {

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

        public DataTable ReadXml(string fileName)
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
        public Form1()
        {
            InitializeComponent();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }



        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "XML|*.xml";
            if (this.saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DataTable dt = (DataTable)this.dataGridView1.DataSource;
                dt.WriteXml(this.saveFileDialog1.FileName, XmlWriteMode.WriteSchema);
            }
        }

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

        private void button3_Click(object sender, EventArgs e)
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

        public void dataGridView1_Load(object sender, EventArgs e)

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

      //  private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
   //     {
          //  for (int i = ; i < 3; i++)
            //{
           //     comboBox1.Items.Add(dataGridView1.Columns[i].HeaderText);

           // }
            
            // var headers = dataGridView1.Columns;

            //  foreach (var header in headers)
            //  {
            //     comboBox1.Items.Add(header.ToString());
            //}
      //  }

        private void cSVToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            DateTime thisDate1 = new DateTime();
            dataGridView1.Columns["dates"].DefaultCellStyle.Format = thisDate1.ToString("MMMM dd, yyyy");
        }
    }
}
