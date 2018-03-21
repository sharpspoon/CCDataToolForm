using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace DataAnalysisTool
{
    public partial class DataAnalysisTool : Form
    {

        //------------------DATE CONVERTER START------------------------------------------------------
        private void dateConvert_Click1(object sender, EventArgs e)
        {

            if (textBox2.Text.Length == 0)
            {
                MessageBox.Show("You did not enter a column name!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                try
                {
                    var value2 = dataGridView1.Rows[i].Cells[textBox2.Text].Value.ToString();
                    if (checkBox1.Checked)
                    {


                        if (value2 == " " || value2 == "" || value2 == null)
                        {
                            MessageBox.Show("NULL at line " + (i + 1) + "\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   NULL at line " + (i + 1) + "\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }
                    }
                    if (value2.Length == 8)
                    {
                        int year = int.Parse(value2.Substring(0, 4));
                        int month = int.Parse(value2.Substring(4, 2));
                        int day = int.Parse(value2.Substring(6, 2));

                        if (year > 2200)
                        {
                            MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month > 12)
                        {
                            MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month < 01)
                        {
                            MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day > 31)
                        {
                            MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day < 01)
                        {
                            MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd");
                        return;
                    }
                }
                catch
                {
                    MessageBox.Show("dates are ok", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   dates are OK");
                    return;
                }



            }


        }
        //------------------DATE CONVERTER END------------------------------------------------------

        //------------------IMPORT FORMAT LOAD START------------------------------------------------------
        private void selectFromDatabaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Importformat importformat = new Importformat();
            importformat.Show();
        }
        private void openImportFormatToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }
        private void openFromZIPExportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string zipPath;
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "ZIP | *.zip", ValidateNames = true, Multiselect = false })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    zipPath = ofd.FileName;
                    string extractPath = @"C:\Program Files (x86)\DataAnalysisTool\ZIP Extracts\" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss");
                    ZipFile.ExtractToDirectory(zipPath, extractPath);
                    MessageBox.Show("Import Format Loaded", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                }
                else
                {
                    MessageBox.Show("error", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        //------------------IMPORT FORMAT LOAD END------------------------------------------------------
        public DataAnalysisTool()
        {
            InitializeComponent();
        }
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = @"TALLYCENTRAL\"+Environment.UserName;
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
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
        private void form1BindingSource_CurrentChanged(object sender, EventArgs e)
        {
        }
        private void ssms_Click(object sender, EventArgs e)
        {
            Ssms ssms = new Ssms();
            ssms.Show();
        }

        private void toolStripStatusLabel4_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel4.Text = dataGridView1.Rows.Count.ToString();
        }

        private void checkToolsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CheckTools cu = new CheckTools();
            cu.Show();
        }

        //------------------OPEN/SAVE XLS START------------------------------------------------------

        private void menu_Open_Xls_Click(object sender, EventArgs e)
        {
            try
            {

                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "XLS | *.xls", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        dataGridView1.DataSource = ReadXls(ofd.FileName);
                    toolStripStatusLabel13.Text = ofd.FileName;
                    toolStripStatusLabel4.Text = dataGridView1.Rows.Count.ToString();
                    richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading XLS: " + ofd.FileName + "...Done.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public DataTable ReadXls(string fileName)
        {
            DataTable dt = new DataTable();
            using (OleDbConnection cn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"" +
                Path.GetDirectoryName(fileName) + "\";Extended Properties='Excel 8.0;HDR=YES;';"))
            {
                using (OleDbCommand cmd = new OleDbCommand("select * from [" + fileName + "$]", cn))
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
        //------------------OPEN/SAVE XLS END------------------------------------------------------

        /// -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------FINALIZED CODE START

        //------------------FORM DRAG LOGIC START------------------------------------------------------
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();
        private void Form1_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }
        //------------------FORM DRAG LOGIC END------------------------------------------------------

        //------------------TOOLSTRIP MINIMIZE, MAXIMIZE, CLOSE START------------------------------------------------------
        private void toolStripMenuItemClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void toolStripMenuItemMaximize_Click(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Normal;
            }
            else
            {
                this.WindowState = FormWindowState.Maximized;
            }
        }
        private void toolStripMenuItemMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        //------------------TOOLSTRIP MINIMIZE, MAXIMIZE, CLOSE END------------------------------------------------------

        //------------------CUT, COPY, PASTE START------------------------------------------------------
        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Control ctrl = this.ActiveControl;
            if (ctrl != null)
            {
                if (ctrl is TextBox)
                {
                    TextBox tx = (TextBox)ctrl;
                    tx.Copy();
                }
            }
        }
        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Control ctrl = this.ActiveControl;
            if (ctrl != null)
            {
                if (ctrl is TextBox)
                {
                    TextBox tx = (TextBox)ctrl;
                    tx.Cut();
                }
            }
        }
        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Control ctrl = this.ActiveControl;
            if (ctrl != null)
            {
                if (ctrl is TextBox)
                {
                    TextBox tx = (TextBox)ctrl;
                    tx.Paste();
                }
            }
        }
        //------------------CUT, COPY, PASTE END------------------------------------------------------

        //------------------CROW NUMBER LOGIC START------------------------------------------------------
        private void dgvUserDetails_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e) 
        {
            using (SolidBrush b = new SolidBrush(dataGridView1.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }
        //------------------CROW NUMBER LOGIC END------------------------------------------------------

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

        //------------------CC LOG OPEN START------------------------------------------------------
        private void cCDataToolLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(Application.UserAppDataPath + @"\Logs");
        }
        //------------------CC LOG OPEN END------------------------------------------------------

        //------------------PRINT DOCUMENT START------------------------------------------------------
        Bitmap bitmap;
        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0 || dataGridView1.Rows == null)
            {
                MessageBox.Show("No data to print", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        //------------------OPEN/SAVE CSV START------------------------------------------------------
        private void menu_Open_Csv_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "CSV | *.csv", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                        dataGridView1.DataSource = ReadCsv(ofd.FileName);
                    toolStripStatusLabel13.Text = ofd.FileName;
                    toolStripStatusLabel13.Visible = true;
                    toolStripStatusLabel4.Text = dataGridView1.Rows.Count.ToString();
                    toolStripStatusLabel2.Visible = true;
                    toolStripStatusLabel3.Visible = true;
                    toolStripStatusLabel4.Visible = true;
                    toolStripStatusLabel5.Visible = true;
                    toolStripStatusLabel12.Visible = true;
                    richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading CSV: " + ofd.FileName + "...Done.");
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

                    toolStripStatusLabel13.Text = ofd.FileName;
                    toolStripStatusLabel4.Text = dataGridView1[0, dataGridView1.Rows.Count - 1].Value.ToString();
                    richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading XML: " + ofd.FileName + "...Done.");
                    toolStripStatusLabel2.Visible = true;
                    toolStripStatusLabel3.Visible = true;
                    toolStripStatusLabel4.Visible = true;
                    toolStripStatusLabel5.Visible = true;
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

        //------------------EXIT APP ACTION START------------------------------------------------------
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                //MessageBox.Show(Application.UserAppDataPath);
                DialogResult result = MessageBox.Show("Do you really want to exit?", "Data Analysis Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    notifyIcon1.Visible = false;
                    notifyIcon1.Icon = null;
                    notifyIcon1.Dispose();
                    if (richTextBox1.Text == "")
                        Environment.Exit(0);
                    else
                    {
                        System.IO.Directory.CreateDirectory(Application.UserAppDataPath + @"\Logs");
                        string path = Application.UserAppDataPath + @"\Logs\DataAnalysisTool_Log_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
                        using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                        {
                            using (TextWriter tw = new StreamWriter(fs))
                            {

                                tw.WriteLine("Data Analysis Tool - Activity Log");
                                tw.WriteLine("Log begin...");
                                tw.WriteLine(".");
                                tw.WriteLine(".");
                                tw.WriteLine(".");
                                tw.WriteLine(richTextBox1.Text);
                                tw.WriteLine("EOF.");

                            }
                        }
                        Environment.Exit(0);
                    }
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
        //------------------EXIT APP ACTION END------------------------------------------------------

        /// -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------FINALIZED CODE END
        /// TEST CODE

        /// TEST CODE

    }
}