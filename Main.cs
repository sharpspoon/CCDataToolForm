using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.IO.Compression;

namespace CCDataImportTool
{
    public partial class CCDataTool : Form
    {
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

        //------------------ABOUT START------------------------------------------------------

        private void menu_About_Click(object sender, EventArgs e)
        {
            About about = new About();
            about.Show();
        }

        //------------------ABOUT END------------------------------------------------------

        //------------------ACKTEKSOFT LOGIN START------------------------------------------------------

        private void acteksoft_Click(object sender, EventArgs e)
        {
            acteksoft actek = new acteksoft();
            actek.Show();
        }
        //------------------ACKTEKSOFT LOGIN END------------------------------------------------------

        //------------------CC LOGO CLICK START------------------------------------------------------
        private void ccLogo_Click1(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://calliduscloud.com");
        }

        //------------------CC LOGO CLICK END------------------------------------------------------



        //------------------SQL LOADER START------------------------------------------------------

        private void serverSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc = new SqlCommand("SELECT name FROM [master].[sys].[databases] where name <> 'master' and name <> 'tempdb' and name <> 'model' and name <> 'msdb' and name <> 'DBAtools'", conn);
            SqlDataReader reader;
            try
            {
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);

                //comboBox2.ValueMember = "1";
                databaseSelect.DataSource = dt;
                databaseSelect.DisplayMember = "name";
                conn.Close();
            }
            catch {
                conn.Close();
                return; }
        }

        private void databaseSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            string ID = databaseSelect.SelectedValue.ToString();
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc = new SqlCommand("use " + databaseSelect.Text + " SELECT table_name AS name FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' order by TABLE_NAME", conn);
            SqlDataReader reader;
            try
            {
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);

                //comboBox2.ValueMember = "1";
                tableSelect.DataSource = dt;
                tableSelect.DisplayMember = "name";
                conn.Close();
            }
            catch { return; }

            conn.Close();
        }

        private void tableSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                var select = "USE "+ databaseSelect.Text+" SELECT * FROM " + tableSelect.Text;
                var conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True"); 
                var dataAdapter = new SqlDataAdapter(select, conn);
                var commandBuilder = new SqlCommandBuilder(dataAdapter);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                dataGridView2.ReadOnly = true;
                dataGridView2.DataSource = ds.Tables[0];
                textBox8.Text = dataGridView2.Rows.Count.ToString();
                conn.Close();
            }
            catch {
                return; }
        }

        //------------------SQL LOADER END------------------------------------------------------

        //------------------IMPORT FORMAT LOAD START------------------------------------------------------
        

        private void openImportFormatToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string zipPath;
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "ZIP | *.zip", ValidateNames = true, Multiselect = false })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    zipPath = ofd.FileName;
                    string extractPath = @"C:\Program Files (x86)\CCDataTool\ZIP Extracts\" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss");
                    ZipFile.ExtractToDirectory(zipPath, extractPath);
                    MessageBox.Show("Import Format Loaded", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                }

                else
                {
                    MessageBox.Show("error", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        //------------------IMPORT FORMAT LOAD END------------------------------------------------------


        public CCDataTool()
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
        private void textBox1_TextChanged(object sender, EventArgs e)
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
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
        private void form1BindingSource_CurrentChanged(object sender, EventArgs e)
        {
        }
        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            textBox7.Text = dataGridView1.Rows.Count.ToString();
        }
        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }
        private void groupBox9_Enter(object sender, EventArgs e)
        {
        }
        private void label8_Click(object sender, EventArgs e)
        {
        }

        private void ssms_Click(object sender, EventArgs e)
        {
            Ssms ssms = new Ssms();
            ssms.Show();
        }

        private void dgvUserDetails_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(dataGridView1.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }
    }
}
