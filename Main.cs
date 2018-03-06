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
                    notifyIcon1.Visible = false;
                    notifyIcon1.Icon = null;
                    notifyIcon1.Dispose();
                    System.IO.Directory.CreateDirectory(@"C:\Program Files (x86)\CCDataTool\Logs");
                    string path = @"C:\Program Files (x86)\CCDataTool\Logs\CCDataTool_Log_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
                    using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                    {
                        using (TextWriter tw = new StreamWriter(fs))
                        {

                            tw.WriteLine("CCDataTool - Error Log");
                            tw.WriteLine("Log begin...");
                            tw.WriteLine(".");
                            tw.WriteLine(".");
                            tw.WriteLine(".");
                            tw.WriteLine(richTextBox1.Text);
                            
                        }
                    }
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

            if (textBox2.Text.Length == 0)
            {
                MessageBox.Show("You did not enter a column name!\r\nThe operation will now cancel.", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
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
                            MessageBox.Show("NULL at line " + (i + 1) + "\r\nMake sure that the date is in the format: yyyyMMdd", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
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
                            MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month > 12)
                        {
                            MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month < 01)
                        {
                            MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day > 31)
                        {
                            MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day < 01)
                        {
                            MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd");
                        return;
                    }
                }
                catch
                {
                    MessageBox.Show("dates are ok", "CCDataTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   dates are OK");
                    return;
                }



            }


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
        private void dgvUserDetails_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e) //row number logic
        {
            using (SolidBrush b = new SolidBrush(dataGridView1.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
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
    }
}