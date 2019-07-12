using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Text;
using PgpCore;

namespace SAPDataAnalysisTool
{

    public partial class SAPDataAnalysisTool : Form
    {
        /*
         * ############################################################################################   
         * ############################################################################################
         * ####################PRODUCTION CODE BEGIN###################################################
         * ############################################################################################
         * ############################################################################################
        */

        //*********************************************************************************************
        //*********************************GLOBAL******************************************************
        //*********************************************************************************************
        public SAPDataAnalysisTool()
        {
            InitializeComponent();
            dateComboBox1.SelectedIndex = 12;
            dateComboBox2.SelectedIndex = 5;
            dateComboBox3.SelectedIndex = 1;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            unableToRegUserToolStripStatusLabel.Text = @"TALLYCENTRAL\" + Environment.UserName;
        }

        Loading loading = new Loading();

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

        //------------------CROW NUMBER LOGIC START------------------------------------------------------
        private void dgvUserDetails_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            using (SolidBrush b = new SolidBrush(importedfileDataGridView.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font, b, e.RowBounds.Location.X + 10, e.RowBounds.Location.Y + 4);
            }
        }
        //------------------CROW NUMBER LOGIC END------------------------------------------------------

        private void Form_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        //------------------TOOLTIP LOGIC START------------------------------------------------------

        ToolTip tt = new ToolTip();

        private void serverSelect_MouseHover(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip1 = new System.Windows.Forms.ToolTip();
            ToolTip1.SetToolTip(this.serverSelect, "Select your ICM server.");
        }

        private void databaseSelect_MouseHover(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.databaseSelect, "Select your ICM database.");
        }

        private void ifSelect_MouseHover(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.ifSelect, "Select your Import Format.");

        }

        private void groupBox7_MouseHover(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.importFormatServerSelectGroupBox, "Select your Server/Database/Import Format.");
        }

        private void reqListBox_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.reqListBox, "Select your required Import Format fields.");
        }

        private void groupBox1_MouseHover(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.importFormatSelectRequiredFieldsGroupBox, "Select your required Import Format fields.");
        }

        private void dateListBox_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.dateListBox, "Select the columns your created date format should apply to.");
        }

        private void dateComboBox1_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.dateComboBox1, "Use this dropdown to build your date format.");
        }

        private void dateComboBox2_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.dateComboBox2, "Use this dropdown to build your date format.");
        }

        private void dateComboBox3_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.dateComboBox3, "Use this dropdown to build your date format.");
        }

        private void dateComboBoxSeperator_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.dateComboBoxSeperator, "Do you want to use a seperator?");
        }

        private void dateFormat_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.dateFormat, "This is the current date format you built");
        }

        private void checkBox2_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.importFormatFindNullCheckbox, "Do you want to find NULLs in the date column?");
        }

        private void tableSelect_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.tableSelect, "Use this dropdown to check any table within your selected database.");
        }

        //------------------TOOLTIP LOGIC END------------------------------------------------------

        private void toolStripStatusLabel18_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.sap.com/index.html");
        }

        private void toolStripStatusLabel19_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.calliduscloud.com/");
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
    (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        //*********************************************************************************************
        //*********************************/GLOBAL*****************************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************HEADER MENU*************************************************
        //*********************************************************************************************

        //------------------CC LOG OPEN START------------------------------------------------------
        private void cCDataToolLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            try
            {
                Process.Start(Application.UserAppDataPath + @"\Logs");
                progressBar1.MarqueeAnimationSpeed = 0;
            }
            catch
            {
                progressBar1.MarqueeAnimationSpeed = 0;
            }
        }
        //------------------CC LOG OPEN END------------------------------------------------------

        //------------------OPEN/SAVE XML START------------------------------------------------------
        private void menu_Open_Xml_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            try
            {
                DataSet dataSet = new DataSet();
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "XML | *.xml", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        dataSet.ReadXml(ofd.FileName);
                        importedfileDataGridView.DataSource = dataSet.Tables[0];

                        importFormatActualFileNameToolStripStatusLabel.Text = ofd.FileName;
                        ifRowCounterToolStripStatusLabel.Text = importedfileDataGridView[0, importedfileDataGridView.Rows.Count - 1].Value.ToString();
                        systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading XML: " + ofd.FileName + "...Done.");
                        ifRowCountLabelToolStripStatusLabel.Visible = true;
                        ifRowCounterToolStripStatusLabel.Visible = true;
                        seperator3ToolStripStatusLabel.Visible = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.MarqueeAnimationSpeed = 0;
        }
        //------------------OPEN/SAVE XML END------------------------------------------------------

        //------------------OPEN/SAVE XLS START------------------------------------------------------

        private void menu_Open_Xls_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            try
            {
                OpenFileDialog openfile1 = new OpenFileDialog();
                if (openfile1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    this.importFormatActualFileNameToolStripStatusLabel.Text = openfile1.FileName;
                }
                {
                    string pathconn = "Provider = Microsoft.jet.OLEDB.4.0; Data source=" + importFormatActualFileNameToolStripStatusLabel.Text + ";Extended Properties=\"Excel 8.0;HDR= yes;\";";
                    OleDbConnection conn = new OleDbConnection(pathconn);
                    OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from [Sheet1$]", conn);
                    DataTable dt = new DataTable();
                    MyDataAdapter.Fill(dt);
                    importedfileDataGridView.DataSource = dt;
                }
            }
            catch { }
            progressBar1.MarqueeAnimationSpeed = 0;
        }
        //------------------OPEN/SAVE XLS END------------------------------------------------------

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

        //------------------TOOLSTRIP MINIMIZE, MAXIMIZE, CLOSE START------------------------------------------------------
        private void toolStripMenuItemClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void toolStripMenuItemMaximize_Click(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
                this.WindowState = FormWindowState.Normal;
            }
            else
            {
                this.MaximizedBounds = Screen.FromHandle(this.Handle).WorkingArea;
                this.WindowState = FormWindowState.Maximized;
            }
        }
        private void toolStripMenuItemMinimize_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        //------------------TOOLSTRIP MINIMIZE, MAXIMIZE, CLOSE END------------------------------------------------------

        //------------------PRINT DOCUMENT START------------------------------------------------------
        Bitmap bitmap;
        private void btnPrint_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            if (importedfileDataGridView.Rows.Count == 0 || importedfileDataGridView.Rows == null)
            {
                MessageBox.Show("No data to print", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                //Resize DataGridView to full height.
                int height = importedfileDataGridView.Height;
                importedfileDataGridView.Height = importedfileDataGridView.RowCount * importedfileDataGridView.RowTemplate.Height;

                //Create a Bitmap and draw the DataGridView on it.
                bitmap = new Bitmap(this.importedfileDataGridView.Width, this.importedfileDataGridView.Height);
                importedfileDataGridView.DrawToBitmap(bitmap, new Rectangle(0, 0, this.importedfileDataGridView.Width, this.importedfileDataGridView.Height));

                //Resize DataGridView back to original height.
                importedfileDataGridView.Height = height;

                //Show the Print Preview Dialog.
                printPreviewDialog1.Document = printDocument1;
                printPreviewDialog1.PrintPreviewControl.Zoom = 1;
                printPreviewDialog1.ShowDialog();
            }
            progressBar1.MarqueeAnimationSpeed = 0;
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
            progressBar1.MarqueeAnimationSpeed = 1;

            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "CSV | *.csv", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        importedfileDataGridView.DataSource = ReadCsv(ofd.FileName);
                        importFormatActualFileNameToolStripStatusLabel.Text = ofd.FileName;
                        importFormatActualFileNameToolStripStatusLabel.Visible = true;
                        ifRowCounterToolStripStatusLabel.Text = importedfileDataGridView.Rows.Count.ToString();
                        ifRowCountLabelToolStripStatusLabel.Visible = true;
                        ifRowCounterToolStripStatusLabel.Visible = true;
                        seperator3ToolStripStatusLabel.Visible = true;
                        importFormatFileNameToolStripStatusLabel.Visible = true;
                        systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading CSV: " + ofd.FileName + "...Done.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            var importedFileArray = importedfileDataGridView.Columns.Cast<DataGridViewColumn>()
                .Select(x => x.HeaderCell.Value.ToString().Trim()).ToArray();



            dateCheckerListBox.Items.Clear();
            specialCharacterCheckerListBox.Items.Clear();
            nullCheckerListBox.Items.Clear();
            cellLengthCheckerListBox.Items.Clear();
            int a = 0;
            for (int i = 0; i < importedFileArray.Length; i++)
            {
                a++;

                specialCharacterCheckerListBox.Items.Add(a + ". " + importedFileArray[i].ToString());
                dateCheckerListBox.Items.Add(a + ". " + importedFileArray[i].ToString());
                nullCheckerListBox.Items.Add(a + ". " + importedFileArray[i].ToString());
                cellLengthCheckerListBox.Items.Add(a + ". " + importedFileArray[i].ToString());
            }

            progressBar1.MarqueeAnimationSpeed = 0;
        }
        public DataTable ReadCsv(string fileName)
        {
            importFormatProgressBar.Value = 0;
            importFormatProgressBar.Value = 20;
            System.Threading.Thread.Sleep(50);
            importFormatProgressBar.Value = 40;
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
            importFormatProgressBar.Value = 100;
            return dt;
        }
        //------------------OPEN/SAVE CSV END------------------------------------------------------

        //------------------ABOUT START------------------------------------------------------
        private void menu_About_Click(object sender, EventArgs e)
        {
            About about = new About();
            about.Show();
        }
        //------------------ABOUT END------------------------------------------------------

        //------------------EXIT APP ACTION START------------------------------------------------------
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            if (e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult result = MessageBox.Show("Do you really want to exit?", "Data Analysis Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    notifyIcon1.Visible = false;
                    notifyIcon1.Icon = null;
                    notifyIcon1.Dispose();
                    if (systemLogTextBox.Text == "")
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
                                tw.WriteLine(systemLogTextBox.Text);
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
            progressBar1.MarqueeAnimationSpeed = 0;
        }
        //------------------EXIT APP ACTION END------------------------------------------------------

        //------------------ACKTEKSOFT LOGIN START------------------------------------------------------
        private void acteksoft_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 10;

            acteksoft actek = new acteksoft();

            while (Application.OpenForms.Count > 1)
            {
                Application.OpenForms[Application.OpenForms.Count - 1].Close();
            }
            actek.ShowDialog();
            progressBar1.MarqueeAnimationSpeed = 0;
        }
        //------------------ACKTEKSOFT LOGIN END------------------------------------------------------


        //*********************************************************************************************
        //*********************************/HEADER MENU************************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************IMPORT FORMAT TAB*******************************************
        //*********************************************************************************************

        private void button7_Click(object sender, EventArgs e)
        {
            Process.Start(Application.UserAppDataPath + @"\IF_Error_Files");
        }

        private void tXTToolStripMenuItemComma_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "TXT | *.txt", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        importedfileDataGridView.DataSource = ReadTxtComma(ofd.FileName);
                        importFormatActualFileNameToolStripStatusLabel.Text = ofd.FileName;
                        importFormatActualFileNameToolStripStatusLabel.Visible = true;
                        ifRowCounterToolStripStatusLabel.Text = importedfileDataGridView.Rows.Count.ToString();
                        ifRowCountLabelToolStripStatusLabel.Visible = true;
                        ifRowCounterToolStripStatusLabel.Visible = true;
                        seperator3ToolStripStatusLabel.Visible = true;
                        importFormatFileNameToolStripStatusLabel.Visible = true;
                        systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading TXT: " + ofd.FileName + "...Done.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.MarqueeAnimationSpeed = 0;
        }

        //*********************************************************************************************
        //*********************************/IMPORT FORMAT TAB******************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************CHECK TOOLS TAB*********************************************
        //*********************************************************************************************

        //------------------SELECT/CLEAR LIST BOX START------------------------------------------------------
        

        //------------------SELECT/CLEAR LIST BOX END------------------------------------------------------

        //------------------DATE CONVERTER START------------------------------------------------------

        //------------------DATE CONVERTER END------------------------------------------------------

        //------------------NULL CHECKER START------------------------------------------------------
        
        //------------------NULL CHECKER END------------------------------------------------------

        //------------------CELL LENGTH CHECKER START------------------------------------------------------
        

        //------------------CELL LENGTH CHECKER END------------------------------------------------------

        //------------------SPECIAL CHARACTER CHECKER START------------------------------------------------------
        

        //------------------SPECIAL CHARACTER CHECKER END------------------------------------------------------

        //*********************************************************************************************
        //*********************************/CHECK TOOLS TAB********************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************SQL QUERY TAB**********************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************/SQL QUERY TAB**********************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************PAYOUT BENCHMARK TAB****************************************
        //*********************************************************************************************

        private void pendingRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (payoutTypeSelect.Text != "")
            {
                int value = payoutTypeSelect.SelectedIndex;
                payoutTypeSelect.SelectedIndex = -1;
                payoutTypeSelect.SelectedIndex = value;
            }
        }

        private void finalizedRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (payoutTypeSelect.Text != "")
            {

                int value = payoutTypeSelect.SelectedIndex;
                payoutTypeSelect.SelectedIndex = -1;
                payoutTypeSelect.SelectedIndex = value;
            }
        }

        private void reversedRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (payoutTypeSelect.Text != "")
            {

                int value = payoutTypeSelect.SelectedIndex;
                payoutTypeSelect.SelectedIndex = -1;
                payoutTypeSelect.SelectedIndex = value;
            }
        }

        private void payoutBenchmarkButton_Click(object sender, EventArgs e)
        {

            benchmarkProgressBar.Value = 0;
            benchmarkProgressBar.Value = 10;

            //global vars
            progressBar1.MarqueeAnimationSpeed = 1;
            if (serverSelect4.Text == "")

            {
                DialogResult result = MessageBox.Show("No server selected. \nPlease make sure you are connected to ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                benchmarkProgressBar.Value = 0;
                return;
            }

            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect4.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();

            //runlistnoroot
            var runListNoRoot = "";
            if (pendingRadioButton.Checked == true)
            {
                runListNoRoot = " USE " + databaseSelect4.Text + " select distinct rl.runlistnoroot from RunList rl left join rundet rd on rd.runlistno = rl.runlistno where rl.rectype='pay' and rd.ItemName = 'PayoutTypeNo' and rd.ItemValue = (select payouttypeno from PayoutType where payouttypeid = '" + payoutTypeSelect.Text + "') and rl.DatFrom = '" + payoutSelect.Text + "' and rl.finalizestatus='p'";
            }
            else if (finalizedRadioButton.Checked == true)
            {
                runListNoRoot = " USE " + databaseSelect4.Text + " select distinct rl.runlistnoroot from RunList rl left join rundet rd on rd.runlistno = rl.runlistno where rl.rectype='pay' and rd.ItemName = 'PayoutTypeNo' and rd.ItemValue = (select payouttypeno from PayoutType where payouttypeid = '" + payoutTypeSelect.Text + "') and rl.DatFrom = '" + payoutSelect.Text + "' and rl.finalizestatus='f'";
            }
            else if (reversedRadioButton.Checked == true)
            {
                runListNoRoot = " USE " + databaseSelect4.Text + " select distinct rl.runlistnoroot from RunList rl left join rundet rd on rd.runlistno = rl.runlistno where rl.rectype='pay' and rd.ItemName = 'PayoutTypeNo' and rd.ItemValue = (select payouttypeno from PayoutType where payouttypeid = '" + payoutTypeSelect.Text + "') and rl.DatFrom = '" + payoutSelect.Text + "' and rl.finalizestatus='r'";
            }
            var dataAdapter3 = new SqlDataAdapter(runListNoRoot, conn);
            var ds3 = new DataSet();
            dataAdapter3.Fill(ds3);
            stagedDataGridView.DataSource = ds3.Tables[0];
            var runListNo = stagedDataGridView.Rows[0].Cells[0].Value;

            //elapsed time
            var elapsedTime = " USE " + databaseSelect4.Text + " select distinct (elapsedtime / 1000) / 60 as name from RunList  where RunListNo = " + runListNo;
            var dataAdapter4 = new SqlDataAdapter(elapsedTime, conn);
            var ds4 = new DataSet();
            dataAdapter4.Fill(ds4);
            stagedDataGridView.DataSource = ds4.Tables[0];
            var elapsedTimeActual = stagedDataGridView.Rows[0].Cells[0].Value;

            //elapsed time average
            var elapsedTimeAverage = "";
            if (pendingRadioButton.Checked == true)
            {
                elapsedTimeAverage = " USE " + databaseSelect4.Text + " select ((sum(elapsedtime)/COUNT(*))/1000) / 60 as name from RunList rl left join rundet rd on rd.runlistno = rl.runlistno where rl.rectype='pay' and rd.ItemName = 'PayoutTypeNo' and rd.ItemValue = (select payouttypeno from PayoutType where payouttypeid = '" + payoutTypeSelect.Text + "')  and rl.finalizestatus='p'";
            }
            else if (finalizedRadioButton.Checked == true)
            {
                elapsedTimeAverage = " USE " + databaseSelect4.Text + " select ((sum(elapsedtime)/COUNT(*))/1000) / 60 as name from RunList rl left join rundet rd on rd.runlistno = rl.runlistno where rl.rectype='pay' and rd.ItemName = 'PayoutTypeNo' and rd.ItemValue = (select payouttypeno from PayoutType where payouttypeid = '" + payoutTypeSelect.Text + "')  and rl.finalizestatus='f'";
            }
            else if (reversedRadioButton.Checked == true)
            {
                elapsedTimeAverage = " USE " + databaseSelect4.Text + " select ((sum(elapsedtime)/COUNT(*))/1000) / 60 as name from RunList rl left join rundet rd on rd.runlistno = rl.runlistno where rl.rectype='pay' and rd.ItemName = 'PayoutTypeNo' and rd.ItemValue = (select payouttypeno from PayoutType where payouttypeid = '" + payoutTypeSelect.Text + "')  and rl.finalizestatus='r'";
            }
            var dataAdapter5 = new SqlDataAdapter(elapsedTimeAverage, conn);
            var ds5 = new DataSet();
            dataAdapter5.Fill(ds5);
            stagedDataGridView.DataSource = ds5.Tables[0];
            var elapsedTimeAverageActual = stagedDataGridView.Rows[0].Cells[0].Value;

            //fasterslower
            var fasterSlower = "";
            if (Convert.ToInt32(elapsedTimeActual) < Convert.ToInt32(elapsedTimeAverageActual))
            {
                fasterSlower = "faster";
            }
            else
            {
                fasterSlower = "slower";
            }

            //fasterslowerpercent
            decimal fasterSlowerPercent = 0;
            if (Convert.ToInt32(elapsedTimeActual) < Convert.ToInt32(elapsedTimeAverageActual))
            {
                fasterSlowerPercent = ((Convert.ToDecimal(elapsedTimeAverageActual) / Convert.ToDecimal(elapsedTimeActual)) - 1) * 100;
            }
            else
            {
                fasterSlowerPercent = fasterSlowerPercent = ((Convert.ToDecimal(elapsedTimeActual) / Convert.ToDecimal(elapsedTimeAverageActual)) - 1) * 100;
            }

            //task numbers
            var taskNumber = " USE " + databaseSelect4.Text + " select taskindex+1 as TaskNumber from runlist where RunListNoRoot=" + runListNo + " and TaskId is not null order by elapsedtime desc";
            var dataAdapter6 = new SqlDataAdapter(taskNumber, conn);
            var ds6 = new DataSet();
            dataAdapter6.Fill(ds6);
            stagedDataGridView.DataSource = ds6.Tables[0];
            var taskNumberArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

            //task ids
            var taskIds = " USE " + databaseSelect4.Text + " select taskid from runlist where RunListNoRoot=" + runListNo + " and TaskId is not null order by elapsedtime desc";
            var dataAdapter7 = new SqlDataAdapter(taskIds, conn);
            var ds7 = new DataSet();
            dataAdapter7.Fill(ds7);
            stagedDataGridView.DataSource = ds7.Tables[0];
            var taskIdsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();



            var tasks = " USE " + databaseSelect4.Text + " select taskindex+1 as 'Task #',TaskId as 'Task Name',((sum(elapsedtime)/COUNT(*))/1000) / 60 as 'Task Run Time in Minutes' from runlist where RunListNoRoot=" + runListNo + " and TaskId is not null group by taskid, TaskIndex, ElapsedTime order by elapsedtime desc";
            var dataAdapter8 = new SqlDataAdapter(tasks, conn);
            var ds8 = new DataSet();
            dataAdapter8.Fill(ds8);
            benchmarkDataGridView.DataSource = ds8.Tables[0];

            benchmarkRichTextBox.Text = benchmarkRichTextBox.Text.Insert(0, Environment.NewLine +
                @"###########################################################################################" + System.Environment.NewLine +
                @"########################DataAnalysisTool - Payout Benchmark################################" + System.Environment.NewLine +
                @"###########################################################################################" + System.Environment.NewLine +
                @"Current Date: " + DateTime.Now + System.Environment.NewLine +
                @"Server: " + serverSelect4.Text + System.Environment.NewLine +
                @"Database: " + databaseSelect4.Text + System.Environment.NewLine +
                @"Payout Type: " + payoutTypeSelect.Text + System.Environment.NewLine +
                @"RunListNoRoot: " + runListNo +
                @"" + System.Environment.NewLine +
                @"" + System.Environment.NewLine +
                @"****************************************************" + System.Environment.NewLine +
                @"********************PAYOUT STATS********************" + System.Environment.NewLine +
                @"****************************************************" + System.Environment.NewLine +
                @"Elapsed time: " + elapsedTimeActual + " Minutes" + System.Environment.NewLine +
                @"Average payout time for the " + payoutTypeSelect.Text + " payout: " + elapsedTimeAverageActual + " Minutes" + System.Environment.NewLine +
                @"Percent " + fasterSlower + " than the payout average: " + fasterSlowerPercent + "%" + System.Environment.NewLine +
                @"" + System.Environment.NewLine +
                @"" + System.Environment.NewLine
                );
            progressBar1.MarqueeAnimationSpeed = 0;
            benchmarkProgressBar.Value = 100;
        }

        private void button28_Click(object sender, EventArgs e)
        {
            Process.Start(Application.UserAppDataPath + @"\Payout_Benchmarks");
        }

        private void benchmarkClearResults_Click(object sender, EventArgs e)
        {
            benchmarkRichTextBox.Clear();
        }

        private void benchmarkExportResults_Click(object sender, EventArgs e)
        {
            if (benchmarkRichTextBox.Text == null || benchmarkRichTextBox.Text == "")
            {
                MessageBox.Show("There are no results to export!", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            System.IO.Directory.CreateDirectory(Application.UserAppDataPath + @"\Payout_Benchmarks");
            string path = Application.UserAppDataPath + @"\Payout_Benchmarks\DataAnalysisTool_PB_Data_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (TextWriter tw = new StreamWriter(fs))
                {
                    for (int i = 0; i < benchmarkRichTextBox.Lines.Length; i++)
                    {
                        tw.WriteLine(benchmarkRichTextBox.Lines[i]);
                    }
                    // setup for export
                    benchmarkDataGridView.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                    benchmarkDataGridView.SelectAll();
                    // hiding row headers to avoid extra \t in exported text
                    var rowHeaders = benchmarkDataGridView.RowHeadersVisible;
                    benchmarkDataGridView.RowHeadersVisible = false;

                    // ! creating text from grid values
                    string content = benchmarkDataGridView.GetClipboardContent().GetText();

                    // restoring grid state
                    benchmarkDataGridView.ClearSelection();
                    benchmarkDataGridView.RowHeadersVisible = rowHeaders;
                    tw.WriteLine(content);
                    tw.WriteLine("EOF.");
                }
            }
            importFormatProgressBar.Value = 90;
            importFormatProgressBar.Value = 100;
            MessageBox.Show("Payout Benchmark file has been created. \nLocation: " + path, "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            progressBar1.MarqueeAnimationSpeed = 0;
            Process.Start(path);
        }

        //*********************************************************************************************
        //*********************************/PAYOUT BENCHMARK TAB****************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************API READINESS TAB*******************************************
        //*********************************************************************************************

        private void apiReadinessCheckButton_Click(object sender, EventArgs e)
        {

            apiReadinessProgressBar.Value = 0;
            apiReadinessProgressBar.Value = 10;

            //global vars
            progressBar1.MarqueeAnimationSpeed = 10;
            if (databaseSelect5.Text == "")
            {
                DialogResult result = MessageBox.Show("No database selected. \nPlease make sure you are connected to ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                importFormatProgressBar.Value = 0;
                return;
            }

            if (databaseSelect5.Text != "")
            {

                DialogResult result2 = MessageBox.Show("The DAT will check against the " + databaseSelect5.Text + " database.\nContinue?", "Data Analysis Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                if (result2 == DialogResult.No)
                {
                    progressBar1.MarqueeAnimationSpeed = 0;
                    importFormatProgressBar.Value = 0;
                    return;
                }
            }

            apiUsersComboBox.Items.Clear();
            apiCallButton.Visible = false;
            apiUsersComboBox.Visible = false;
            apiUsersPictureBox.Visible = false;
            apiUsersPasswordPictureBox.Visible = false;
            apiUsersPasswordTextBox.Visible = false;
            apiRichTextBox.Clear();

            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect5.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();

            var secGroups = " USE " + databaseSelect5.Text + " select SecGroupId from secgroup where portalid=6 and prosta=1";
            var dataAdapter = new SqlDataAdapter(secGroups, conn);
            var ds = new DataSet();
            dataAdapter.Fill(ds);
            stagedDataGridView.DataSource = ds.Tables[0];
            var secGroupsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                    .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

            var apiUsers = " USE " + databaseSelect5.Text + " select us.userid from UsrPortal up inner join UsrSet us on up.userno=us.userno where up.ProSta=1 and up.SecGroupNo in (select SecGroupNo from secgroup where portalid=6 and prosta=1)";
            var dataAdapter2 = new SqlDataAdapter(apiUsers, conn);
            var ds2 = new DataSet();
            dataAdapter2.Fill(ds2);
            stagedDataGridView.DataSource = ds2.Tables[0];
            var apiUsersArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                    .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

            var apiEnabled = " USE " + databaseSelect5.Text + " select case when enabled=1 then 'Yes' else 'No' end as 'Enabled' from feature where FeatureId='System API''s'";
            var dataAdapter3 = new SqlDataAdapter(apiEnabled, conn);
            var ds3 = new DataSet();
            dataAdapter3.Fill(ds3);
            stagedDataGridView.DataSource = ds3.Tables[0];
            var apiEnabledFinal = stagedDataGridView.Rows[0].Cells[0].Value;

            conn.Close();

            progressBar1.MarqueeAnimationSpeed = 0;

            apiRichTextBox.AppendText(Environment.NewLine +
                @"###########################################################################################" + System.Environment.NewLine +
                @"########################DataAnalysisTool - API Readiness###################################" + System.Environment.NewLine +
                @"###########################################################################################" + System.Environment.NewLine +
                @"Current Date: " + DateTime.Now + System.Environment.NewLine +
                @"Server: " + serverSelect5.Text + System.Environment.NewLine +
                @"Database: " + databaseSelect5.Text + System.Environment.NewLine +
                @"" + System.Environment.NewLine +
                @"" + System.Environment.NewLine +
                @"****************************************************" + System.Environment.NewLine +
                @"********************RUN RESULTS*********************" + System.Environment.NewLine +
                @"****************************************************" + System.Environment.NewLine
                );

            apiRichTextBox.AppendText(@"" + System.Environment.NewLine + "API enabled: " + System.Environment.NewLine + apiEnabledFinal);

            apiRichTextBox.AppendText(@"" + System.Environment.NewLine);

            if (apiEnabledFinal.Equals("Yes"))
            {
                apiPictureBox.Image = Properties.Resources.greenCheck;
            }
            else
            {
                apiRichTextBox.AppendText(Environment.NewLine + @"Please enable API's within the Global Features.");
                apiPictureBox.Image = Properties.Resources.global;
                return;
            }

            if (secGroupsArray.Length == 0)
            {
                apiRichTextBox.AppendText(Environment.NewLine + @"Please enable or create an API security group.");
                apiPictureBox.Image = Properties.Resources.sec;
                return;
            }
            else
            {
                apiPictureBox.Image = Properties.Resources.greenCheck;
            }

            apiRichTextBox.AppendText(Environment.NewLine + @"API Security Groups:");
            foreach (var sec in secGroupsArray)
            {
                apiRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
            }

            apiRichTextBox.AppendText(Environment.NewLine + @"");

            apiRichTextBox.AppendText(Environment.NewLine + @"Optionally, configure the system.api.ip.whitelist to restrict access to a range of client IP addresses. 
(Admin > Configuration > Options) E.g. restrict access to internal IP addresses. Note that if this option
is not configured, the System API's may be accessed from any IP address. This may be considered a security 
risk if your ICM instance is externally accessible.");

            apiRichTextBox.AppendText(Environment.NewLine + @"");
            apiRichTextBox.AppendText(Environment.NewLine + @"API Users:");

            if (apiUsersArray.Length >= 1)
            {
                foreach (var api in apiUsersArray)
                {
                    apiRichTextBox.AppendText(@"" + System.Environment.NewLine + api);
                }
                apiCallButton.Visible = true;
                apiUsersComboBox.Visible = true;
                apiUsersPictureBox.Visible = true;
                apiUsersPasswordPictureBox.Visible = true;
                apiUsersPasswordTextBox.Visible = true;
                apiEnvLabel1.Visible = true;
                apiEnvLabel2.Visible = true;
                apiEnvLabelMain.Visible = true;
                for (int i = 0; i < apiUsersArray.Length; i++)
                {
                    apiUsersComboBox.Items.Add(apiUsersArray[i]);
                }
                apiRichTextBox.AppendText(Environment.NewLine + @"");
                apiRichTextBox.AppendText(Environment.NewLine + @"It looks like this environment is ready to test an API call. If you would like to do this, please select a user above, type the password, then click Test Call");
                apiCallButton.Visible = true;
            }
            else
            {
                apiRichTextBox.AppendText(Environment.NewLine + @"No API users found. Define one or more Users with access to the '''System API's''' AppArea. (Admin > Security > Users).");
                apiPictureBox.Image = Properties.Resources.apiuser;
                return;
            }
            apiRichTextBox.AppendText(Environment.NewLine + @"");
            apiReadinessProgressBar.Value = 100;
        }

        private void apiCallButton_Click(object sender, EventArgs e)
        {
            apiRichTextBox.Clear();
            using (var client = new HttpClient(new HttpClientHandler { AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate }))
            {
                apiRichTextBox.AppendText(Environment.NewLine +
                    @"###########################################################################################" + System.Environment.NewLine +
                    @"########################DataAnalysisTool - API Readiness###################################" + System.Environment.NewLine +
                    @"###########################################################################################" + System.Environment.NewLine +
                    @"Current Date: " + DateTime.Now + System.Environment.NewLine +
                    @"Server: " + serverSelect5.Text + System.Environment.NewLine +
                    @"Database: " + databaseSelect5.Text + System.Environment.NewLine +
                    @"" + System.Environment.NewLine +
                    @"" + System.Environment.NewLine +
                    @"****************************************************" + System.Environment.NewLine +
                    @"********************API CALL RESULTS****************" + System.Environment.NewLine +
                    @"****************************************************" + System.Environment.NewLine
                    );
                client.BaseAddress = new Uri("https://welltest2.callidusinsurance.net/ICM/REST/auth/login?u=" + apiUsersComboBox.Text + "&p=" + apiUsersPasswordTextBox.Text);
                HttpResponseMessage response = client.GetAsync("").Result;
                response.EnsureSuccessStatusCode();
                string result = response.Content.ReadAsStringAsync().Result;
                Console.WriteLine("Result: " + result);
                apiRichTextBox.AppendText(@"" + System.Environment.NewLine + result);
            }
        }

        private void aPILogFolderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(Application.UserAppDataPath + @"\API_Readiness_Check");
        }

        //*********************************************************************************************
        //*********************************/API READINESS TAB*******************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************PGP TAB*****************************************************
        //*********************************************************************************************
        


        //*********************************************************************************************
        //*********************************/PGP TAB****************************************************
        //*********************************************************************************************


        private void toolStripStatusLabel4_Click(object sender, EventArgs e)
        {
            ifRowCounterToolStripStatusLabel.Text = importedfileDataGridView.Rows.Count.ToString();
        }

        public DataTable ReadTxtComma(string fileName)
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

        private void pipeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "TXT | *.txt", ValidateNames = true, Multiselect = false })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        importFormatProgressBar.Value = 20;
                        importedfileDataGridView.DataSource = ReadTxtPipe(ofd.FileName);
                        importFormatActualFileNameToolStripStatusLabel.Text = ofd.FileName;
                        importFormatActualFileNameToolStripStatusLabel.Visible = true;
                        ifRowCounterToolStripStatusLabel.Text = importedfileDataGridView.Rows.Count.ToString();
                        ifRowCountLabelToolStripStatusLabel.Visible = true;
                        ifRowCounterToolStripStatusLabel.Visible = true;
                        seperator3ToolStripStatusLabel.Visible = true;
                        importFormatFileNameToolStripStatusLabel.Visible = true;
                        systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading TXT: " + ofd.FileName + "...Done.");
                    }
                    else
                    {
                        importFormatProgressBar.Value = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                importFormatProgressBar.Value = 0;
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            importFormatProgressBar.Value = 100;
            progressBar1.MarqueeAnimationSpeed = 0;
        }

        public DataTable ReadTxtPipe(string fileName)
        {
            importFormatProgressBar.Value = 30;
            DataTable dt = new DataTable();
            string[] columns = null;

            var lines = File.ReadAllLines(fileName);

            if (importformatIncludeHeaderRowButton.Checked == false)
            {
                importFormatProgressBar.Value = 50;
                if (lines.Count() > 0)
                {
                    importFormatProgressBar.Value = 60;
                    columns = lines[0].Split(new char[] { '|' });
                }

                int columnCount1 = columns.Count();
                for (int i = 0; i < columnCount1; i++)
                {
                    dt.Columns.Add("column " + (i+1));
                }

                // reading rest of the data
                for (int i = 0; i < lines.Count(); i++)
                {
                    DataRow dr = dt.NewRow();
                    string[] values = lines[i].Split(new char[] { '|' });

                    for (int j = 0; j < values.Count() && j < columns.Count(); j++)
                        dr[j] = values[j];

                    dt.Rows.Add(dr);
                }
                importFormatProgressBar.Value = 70;
                return dt;
            }
            else
            {
                importFormatProgressBar.Value = 50;
                if (lines.Count() > 0)
                {
                    importFormatProgressBar.Value = 60;
                    columns = lines[0].Split(new char[] { '|' });

                    foreach (var column in columns)
                        dt.Columns.Add(column);
                }

                // reading rest of the data
                for (int i = 1; i < lines.Count(); i++)
                {
                    DataRow dr = dt.NewRow();
                    string[] values = lines[i].Split(new char[] { '|' });
                    for (int j = 0; j < values.Count() && j < columns.Count(); j++)
                        dr[j] = values[j];
                    dt.Rows.Add(dr);
                }
                importFormatProgressBar.Value = 70;
                return dt;
            }
        }

        private void dateComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //day check
            if (dateComboBox1.Text == "d" || dateComboBox1.Text == "dd" || dateComboBox1.Text == "ddd" || dateComboBox1.Text == "dddd")
            {
                if (dateComboBox2.Text == "d" || dateComboBox3.Text == "d")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "dd" || dateComboBox3.Text == "dd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "ddd" || dateComboBox3.Text == "ddd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "dddd" || dateComboBox3.Text == "dddd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
            }

            //month check
            if (dateComboBox1.Text == "m" || dateComboBox1.Text == "mm" || dateComboBox1.Text == "M" || dateComboBox1.Text == "MM" || dateComboBox1.Text == "MMM" || dateComboBox1.Text == "MMM" || dateComboBox1.Text == "MMMM")
            {
                if (dateComboBox2.Text == "m" || dateComboBox3.Text == "m")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "mm" || dateComboBox3.Text == "mm")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "M" || dateComboBox3.Text == "M")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "MM" || dateComboBox3.Text == "MM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "MMM" || dateComboBox3.Text == "MMM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "MMMM" || dateComboBox3.Text == "MMMM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
            }

            //year check
            if (dateComboBox1.Text == "y" || dateComboBox1.Text == "yy" || dateComboBox1.Text == "yyyy")
            {
                if (dateComboBox2.Text == "y" || dateComboBox3.Text == "y")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "yy" || dateComboBox3.Text == "yy")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "yyyy" || dateComboBox3.Text == "yyyy")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
            }


            dateFormat.Text = "Date Format: "+dateComboBox1.Text+ dateComboBoxSeperator.Text + dateComboBox2.Text+ dateComboBoxSeperator.Text+dateComboBox3.Text;
        }

        private void dateComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //day check
            if (dateComboBox2.Text == "d" || dateComboBox2.Text == "dd" || dateComboBox2.Text == "ddd" || dateComboBox2.Text == "dddd")
            {
                if (dateComboBox1.Text == "d" || dateComboBox3.Text == "d")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "dd" || dateComboBox3.Text == "dd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "ddd" || dateComboBox3.Text == "ddd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "dddd" || dateComboBox3.Text == "dddd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox1.Text = null;
                    return;
                }
            }

            //month check
            if (dateComboBox2.Text == "m" || dateComboBox2.Text == "mm" || dateComboBox2.Text == "M" || dateComboBox2.Text == "MM" || dateComboBox2.Text == "MMM" || dateComboBox2.Text == "MMM" || dateComboBox2.Text == "MMMM")
            {
                if (dateComboBox1.Text == "m" || dateComboBox3.Text == "m")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "mm" || dateComboBox3.Text == "mm")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "M" || dateComboBox3.Text == "M")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "MM" || dateComboBox3.Text == "MM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "MMM" || dateComboBox3.Text == "MMM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "MMMM" || dateComboBox3.Text == "MMMM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
            }

            //year check
            if (dateComboBox2.Text == "y" || dateComboBox2.Text == "yy" || dateComboBox2.Text == "yyyy")
            {
                if (dateComboBox1.Text == "y" || dateComboBox3.Text == "y")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "yy" || dateComboBox3.Text == "yy")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "yyyy" || dateComboBox3.Text == "yyyy")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox2.Text = null;
                    return;
                }
            }
            dateFormat.Text = "Date Format: " + dateComboBox1.Text + dateComboBoxSeperator.Text + dateComboBox2.Text + dateComboBoxSeperator.Text + dateComboBox3.Text;
        }

        private void dateComboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            //day check
            if (dateComboBox3.Text == "d" || dateComboBox3.Text == "dd" || dateComboBox3.Text == "ddd" || dateComboBox3.Text == "dddd")
            {
                if (dateComboBox1.Text == "d" || dateComboBox2.Text == "d")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "dd" || dateComboBox2.Text == "dd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "ddd" || dateComboBox2.Text == "ddd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "dddd" || dateComboBox2.Text == "dddd")
                {
                    MessageBox.Show("Cannot have more than one 'day' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
            }

            //month check
            if (dateComboBox3.Text == "m" || dateComboBox3.Text == "mm" || dateComboBox3.Text == "M" || dateComboBox3.Text == "MM" || dateComboBox3.Text == "MMM" || dateComboBox3.Text == "MMM" || dateComboBox3.Text == "MMMM")
            {
                if (dateComboBox1.Text == "m" || dateComboBox2.Text == "m")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "mm" || dateComboBox2.Text == "mm")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "M" || dateComboBox2.Text == "M")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "MM" || dateComboBox2.Text == "MM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "MMM" || dateComboBox2.Text == "MMM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox1.Text == "MMMM" || dateComboBox2.Text == "MMMM")
                {
                    MessageBox.Show("Cannot have more than one 'month' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
            }

            //year check
            if (dateComboBox3.Text == "y" || dateComboBox3.Text == "yy" || dateComboBox3.Text == "yyyy")
            {
                if (dateComboBox2.Text == "y" || dateComboBox1.Text == "y")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "yy" || dateComboBox1.Text == "yy")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
                if (dateComboBox2.Text == "yyyy" || dateComboBox1.Text == "yyyy")
                {
                    MessageBox.Show("Cannot have more than one 'year' type", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    dateComboBox3.Text = null;
                    return;
                }
            }
            dateFormat.Text = "Date Format: " + dateComboBox1.Text + dateComboBoxSeperator.Text + dateComboBox2.Text + dateComboBoxSeperator.Text + dateComboBox3.Text;
        }

        private void dateComboBoxSeperator_SelectedIndexChanged(object sender, EventArgs e)
        {
            dateFormat.Text = "Date Format: " + dateComboBox1.Text + dateComboBoxSeperator.Text + dateComboBox2.Text + dateComboBoxSeperator.Text + dateComboBox3.Text;
        }

        private void button25_Click(object sender, EventArgs e)
        {
            try
            {
                int length = int.Parse(importFormatJumpToRowTextBox.Text);
                importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[length - 1].Cells[0];
                importedfileDataGridView.Rows[length - 1].Selected = true;
            }
            catch { MessageBox.Show("That column does not exist!", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1); }
        }

        private void checkBox4_Click(object sender, EventArgs e)
        {
            if (databaseSelect.Text != "")
            {

                int value = databaseSelect.SelectedIndex;
                databaseSelect.SelectedIndex = -1;
                databaseSelect.SelectedIndex = value;
            }
        }

        //------------------EXIT APP ACTION END------------------------------------------------------
        /*
         * ############################################################################################   
         * ############################################################################################
         * ####################PRODUCTION CODE END#####################################################
         * ############################################################################################
         * ############################################################################################
        */

        private void button27_Click(object sender, EventArgs e)
        {
            SqlConnection pubsConn = new SqlConnection(@"Data Source = " + serverSelect5.Text + "; Initial Catalog = master; Integrated Security = True");
            SqlCommand logoCMD = new SqlCommand(" USE " + databaseSelect5.Text + " select content from outfile where runlistno =15408457951330000", pubsConn);

            FileStream fs;                          // Writes the BLOB to a file (*.bmp).
            BinaryWriter bw;                        // Streams the BLOB to the FileStream object.

            int bufferSize = 100;                   // Size of the BLOB buffer.
            byte[] outbyte = new byte[bufferSize];  // The BLOB byte[] buffer to be filled by GetBytes.
            long retval;                            // The bytes returned from GetBytes.
            long startIndex = 0;                    // The starting position in the BLOB output.

            string pub_id = "";                     // The publisher id to use in the file name.

            // Open the connection and read data into the DataReader.
            pubsConn.Open();
            SqlDataReader myReader = logoCMD.ExecuteReader(CommandBehavior.SequentialAccess);

            while (myReader.Read())
            {
                // Get the publisher id, which must occur before getting the logo.
                // Create a file to hold the output.
                fs = new FileStream("icmlog" + pub_id + ".log", FileMode.OpenOrCreate, FileAccess.Write);
                bw = new BinaryWriter(fs);
                // Reset the starting byte for the new BLOB.
                MessageBox.Show("" + startIndex);
                startIndex = 0;
                // Read the bytes into outbyte[] and retain the number of bytes returned.
                MessageBox.Show("outbyte" + outbyte);
                MessageBox.Show("buffersize" + bufferSize);
                retval = myReader.GetBytes(1, startIndex, outbyte, 0, bufferSize);//fails
                MessageBox.Show("1859");
                // Continue reading and writing while there are bytes beyond the size of the buffer.
                while (retval == bufferSize)
                {
                    bw.Write(outbyte);
                    bw.Flush();

                    // Reposition the start index to the end of the last buffer and fill the buffer.
                    startIndex += bufferSize;
                    retval = myReader.GetBytes(1, startIndex, outbyte, 0, bufferSize);
                }

                // Write the remaining buffer.
                bw.Write(outbyte, 0, (int)retval - 1);
                bw.Flush();

                // Close the output file.
                bw.Close();
                fs.Close();
            }

            // Close the reader and the connection.
            myReader.Close();
            pubsConn.Close();
        }



        private void copyAlltoClipboard()
        {
            importedfileDataGridView.SelectAll();
            DataObject dataObj = importedfileDataGridView.GetClipboardContent();
            if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
        }

        private void goButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.goButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.goButtonPictureBox, "Run the tool!");
        }

        private void goButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.goButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.goButtonPictureBox, "Run the tool!");
        }

        private void goButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.goButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go3));
        }

        private void goButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.goButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
        }

        private void csvButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.csvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv3));
        }

        private void csvButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.csvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv2));
        }

        private void csvButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.csvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.csvButtonPictureBox, "Open a CSV file.");
        }

        private void csvButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.csvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv));
        }

        private void xmlButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.xmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml3));
        }

        private void xmlButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.xmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.xmlButtonPictureBox, "Open an XML file.");
        }

        private void xmlButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.xmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.xmlButtonPictureBox, "Open an XML file.");
        }

        private void xmlButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.xmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml));
        }

        private void txtCommaButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.txtCommaButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_comma3));
        }

        private void txtCommaButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.txtCommaButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_comma2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.txtCommaButtonPictureBox, "Open a Text Comma file.");
        }

        private void txtCommaButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.txtCommaButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_comma));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.txtCommaButtonPictureBox, "Open a Text Comma file.");
        }

        private void txtCommaButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.txtCommaButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_comma));
        }

        private void xlsButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.xlsButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xls3));
        }

        private void xlsButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.xlsButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xls2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.xlsButtonPictureBox, "Open an XLS file.");
        }

        private void xlsButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.xlsButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xls));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.xlsButtonPictureBox, "Open an XLS file.");
        }

        private void xlsButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.xlsButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xls));
        }

        private void txtPipePictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.txtPipePictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_pipe3));
        }

        private void txtPipePictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.txtPipePictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_pipe2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.txtPipePictureBox, "Open a Text Pipe file.");
        }

        private void txtPipePictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.txtPipePictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_pipe));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.txtPipePictureBox, "Open a Text Pipe file.");
        }

        private void txtPipePictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.txtPipePictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_txt_pipe));
        }

        private void clearResultsPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.clearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results3));
        }

        private void clearResultsPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.clearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearResultsPictureBox, "Clear the results.");
        }

        private void clearResultsPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.clearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearResultsPictureBox, "Clear the results.");
        }

        private void clearResultsPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.clearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results));
        }

        private void exportResultsPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.exportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results3));
        }

        private void exportResultsPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.exportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.exportResultsPictureBox, "Export the results.");
        }

        private void exportResultsPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.exportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.exportResultsPictureBox, "Export the results.");
        }

        private void exportResultsPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.exportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results));
        }

        private void fromDateEnableCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (fromDateEnableCheckBox.Checked == true)
            {
                dateYearFromTextBox.ReadOnly = false;
                dateMonthFromTextBox.ReadOnly = false;
                dateDayFromTextBox.ReadOnly = false;
            }

            else
            {
                dateYearFromTextBox.ReadOnly = true;
                dateMonthFromTextBox.ReadOnly = true;
                dateDayFromTextBox.ReadOnly = true;
            }
        }

        private void toDateEnableCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (toDateEnableCheckBox.Checked == true)
            {
                dateYearToTextBox.ReadOnly = false;
                dateMonthToTextBox.ReadOnly = false;
                dateDayToTextBox.ReadOnly = false;
            }

            else
            {
                dateYearToTextBox.ReadOnly = true;
                dateMonthToTextBox.ReadOnly = true;
                dateDayToTextBox.ReadOnly = true;
            }
        }

        private void dateTextBox_KeyUp(object sender, KeyEventArgs e)
        {
            dateRangeLabel.Text = "Date Range: " + dateYearFromTextBox.Text + dateMonthFromTextBox.Text + dateDayFromTextBox.Text + " - " + dateYearToTextBox.Text + dateMonthToTextBox.Text + dateDayToTextBox.Text;
        }

        private void envChangesGoPictureBox_Click(object sender, EventArgs e)
        {
            envChangesProgressBar.Value = 0;
            envChangesProgressBar.Value = 10;

            //global vars
            progressBar1.MarqueeAnimationSpeed = 10;


            if (databaseSelect6.Text == "")
            {
                DialogResult result = MessageBox.Show("No database selected. \nPlease make sure you are connected to ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                envChangesProgressBar.Value = 0;
                return;
            }

            if (userIDTextBox.Text == "")
            {
                DialogResult result = MessageBox.Show("No UserID entered. \nPlease enter a UserID", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                envChangesProgressBar.Value = 0;
                return;
            }
            var fromDate = dateYearFromTextBox.Text + dateMonthFromTextBox.Text + dateDayFromTextBox.Text;
            var toDate = dateYearToTextBox.Text + dateMonthToTextBox.Text + dateDayToTextBox.Text;
            if (fromDateEnableCheckBox.Checked == true && fromDate.Length != 8)
            {
                MessageBox.Show("Incorrect date format on the From Section. \nPlease make sure you are using YYYYMMDD", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                envChangesProgressBar.Value = 0;
                return;
            }
            if (toDateEnableCheckBox.Checked == true && toDate.Length != 8)
            {
                MessageBox.Show("Incorrect date format on the From Section. \nPlease make sure you are using YYYYMMDD", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                envChangesProgressBar.Value = 0;
                return;
            }
            envChangesRichTextBox.Clear();
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect6.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            envChangesRichTextBox.AppendText(Environment.NewLine +
                @"###########################################################################################" + System.Environment.NewLine +
                @"########################DataAnalysisTool - Environment Changes#############################" + System.Environment.NewLine +
                @"###########################################################################################" + System.Environment.NewLine +
                @"Current Date: " + DateTime.Now + System.Environment.NewLine +
                @"Server: " + serverSelect6.Text + System.Environment.NewLine +
                @"Database: " + databaseSelect6.Text + System.Environment.NewLine +
                @"User: " + userIDTextBox.Text + System.Environment.NewLine +
                @"" + dateRangeLabel.Text + System.Environment.NewLine +
                @"" + System.Environment.NewLine +
                @"" + System.Environment.NewLine +
                @"****************************************************" + System.Environment.NewLine +
                @"********************RUN RESULTS*********************" + System.Environment.NewLine +
                @"****************************************************" + System.Environment.NewLine
                );

            //Import Formats
            envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine);
            if (envChangesCheckBox1.Checked == true)
            {
                var changedImportformats = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedImportformats = " USE " + databaseSelect6.Text + " select importformatid from importformat where lstuser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedImportformats = " USE " + databaseSelect6.Text + " select importformatid from importformat where lstuser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedImportformats = " USE " + databaseSelect6.Text + " select importformatid from importformat where lstuser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedImportformats = " USE " + databaseSelect6.Text + " select importformatid from importformat where lstuser=" + "'" + userIDTextBox.Text + "'";
                }
                var dataAdapter = new SqlDataAdapter(changedImportformats, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedImportFormatsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed Import Formats:");
                foreach (var sec in changedImportFormatsArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            //Expressions
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            if (envChangesCheckBox2.Checked == true)
            {
                var changedExpressions = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedExpressions = " USE " + databaseSelect6.Text + " select expressionid from expression where lstuser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedExpressions = " USE " + databaseSelect6.Text + " select expressionid from expression where lstuser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedExpressions = " USE " + databaseSelect6.Text + " select expressionid from expression where lstuser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedExpressions = " USE " + databaseSelect6.Text + " select expressionid from expression where lstuser=" + "'" + userIDTextBox.Text + "'";
                }
                var dataAdapter = new SqlDataAdapter(changedExpressions, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedExpressionsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed Expressions:");
                foreach (var sec in changedExpressionsArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            //QBQ
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            if (envChangesCheckBox3.Checked == true)
            {
                var changedQBQ = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedQBQ = " USE " + databaseSelect6.Text + " select QBQueryId from QBQuery where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >"+fromDate+" and lstdate < "+toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedQBQ = " USE " + databaseSelect6.Text + " select QBQueryId from QBQuery where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedQBQ = " USE " + databaseSelect6.Text + " select QBQueryId from QBQuery where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedQBQ = " USE " + databaseSelect6.Text + " select QBQueryId from QBQuery where LstUser=" + "'" + userIDTextBox.Text + "'";
                }
                
                var dataAdapter = new SqlDataAdapter(changedQBQ, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedQBQArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed QBQ:");
                foreach (var sec in changedQBQArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            //Xref
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            if (envChangesCheckBox4.Checked == true)
            {
                var changedXref = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedXref = " USE " + databaseSelect6.Text + " select ExtCrossRefTypeId from ExtCrossRefType where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedXref = " USE " + databaseSelect6.Text + " select ExtCrossRefTypeId from ExtCrossRefType where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedXref = " USE " + databaseSelect6.Text + " select ExtCrossRefTypeId from ExtCrossRefType where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedXref = " USE " + databaseSelect6.Text + " select ExtCrossRefTypeId from ExtCrossRefType where LstUser=" + "'" + userIDTextBox.Text + "'";
                }
                var dataAdapter = new SqlDataAdapter(changedXref, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedImportFormatsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed Cross-Refs:");
                foreach (var sec in changedImportFormatsArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            //Field Default
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            if (envChangesCheckBox5.Checked == true)
            {
                var changedFieldDefault = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedFieldDefault = " USE " + databaseSelect6.Text + " select 'EntName: '+EntName+' FldName: '+FldName from FieldDefault where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedFieldDefault = " USE " + databaseSelect6.Text + " select 'EntName: '+EntName+' FldName: '+FldName from FieldDefault where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedFieldDefault = " USE " + databaseSelect6.Text + " select 'EntName: '+EntName+' FldName: '+FldName from FieldDefault where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedFieldDefault = " USE " + databaseSelect6.Text + " select 'EntName: '+EntName+' FldName: '+FldName from FieldDefault where LstUser=" + "'" + userIDTextBox.Text + "'";
                }
                var dataAdapter = new SqlDataAdapter(changedFieldDefault, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedImportFormatsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed Field Defaults:");
                foreach (var sec in changedImportFormatsArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            //BEU
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            if (envChangesCheckBox6.Checked == true)
            {
                var changedBEU = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedBEU = " USE " + databaseSelect6.Text + " select BatchEntityUpdateId from BatchEntityUpdate where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedBEU = " USE " + databaseSelect6.Text + " select BatchEntityUpdateId from BatchEntityUpdate where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedBEU = " USE " + databaseSelect6.Text + " select BatchEntityUpdateId from BatchEntityUpdate where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedBEU = " USE " + databaseSelect6.Text + " select BatchEntityUpdateId from BatchEntityUpdate where LstUser=" + "'" + userIDTextBox.Text + "'";
                }
                var dataAdapter = new SqlDataAdapter(changedBEU, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedImportFormatsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed BEUs:");
                foreach (var sec in changedImportFormatsArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            //Report Forms
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            if (envChangesCheckBox6.Checked == true)
            {
                var changedReportForms = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedReportForms = " USE " + databaseSelect6.Text + " select Name from ReportForm where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedReportForms = " USE " + databaseSelect6.Text + " select Name from ReportForm where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedReportForms = " USE " + databaseSelect6.Text + " select Name from ReportForm where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedReportForms = " USE " + databaseSelect6.Text + " select Name from ReportForm where LstUser=" + "'" + userIDTextBox.Text + "'";
                }
                var dataAdapter = new SqlDataAdapter(changedReportForms, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedImportFormatsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed Report Forms:");
                foreach (var sec in changedImportFormatsArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            //Report Templates
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            if (envChangesCheckBox6.Checked == true)
            {
                var changedReportTemplates = "";
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8 && toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedReportTemplates = " USE " + databaseSelect6.Text + " select ReportId from JasperReport where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == true && fromDate.Length == 8)
                {
                    changedReportTemplates = " USE " + databaseSelect6.Text + " select ReportId from JasperReport where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate >" + fromDate;
                }
                if (toDateEnableCheckBox.Checked == true && toDate.Length == 8)
                {
                    changedReportTemplates = " USE " + databaseSelect6.Text + " select ReportId from JasperReport where LstUser=" + "'" + userIDTextBox.Text + "'" + " and lstdate < " + toDate;
                }
                if (fromDateEnableCheckBox.Checked == false && toDateEnableCheckBox.Checked == false)
                {
                    changedReportTemplates = " USE " + databaseSelect6.Text + " select ReportId from JasperReport where LstUser=" + "'" + userIDTextBox.Text + "'";
                }
                var dataAdapter = new SqlDataAdapter(changedReportTemplates, conn);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                stagedDataGridView.DataSource = ds.Tables[0];
                var changedImportFormatsArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                envChangesRichTextBox.AppendText(Environment.NewLine + @"Changed Report Templates:");
                foreach (var sec in changedImportFormatsArray)
                {
                    envChangesRichTextBox.AppendText(@"" + System.Environment.NewLine + sec);
                }
            }
            envChangesRichTextBox.AppendText(Environment.NewLine + @"");
            envChangesProgressBar.Value = 100;
        }

        private void envChangesGoPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.envChangesGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go3));
        }

        private void envChangesGoPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.envChangesGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.envChangesGoPictureBox, "Run the tool!");
        }

        private void envChangesGoPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.envChangesGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.envChangesGoPictureBox, "Run the tool!");
        }

        private void envChangesGoPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.envChangesGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
        }

        private void clearResultsPictureBox_Click(object sender, EventArgs e)
        {
            envChangesRichTextBox.Clear();
        }

        private void apiGoPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.apiGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go3));
        }

        private void apiGoPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.apiGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.apiGoPictureBox, "Run the tool!");
        }

        private void apiGoPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.apiGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.apiGoPictureBox, "Run the tool!");
        }

        private void apiGoPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.apiGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
        }

        private void apiExportResultsPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.apiExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results3));
        }

        private void apiExportResultsPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.apiExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.apiExportResultsPictureBox, "Export the results.");
        }

        private void apiExportResultsPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.apiExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.apiExportResultsPictureBox, "Export the results.");
        }

        private void apiExportResultsPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.apiExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results));
        }

        private void apiClearResultsPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.apiClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results3));
        }

        private void apiClearResultsPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.apiClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.apiClearResultsPictureBox, "Clear the results.");
        }

        private void apiClearResultsPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.apiClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.apiClearResultsPictureBox, "Clear the results.");
        }

        private void apiClearResultsPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.apiClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results));
        }

        private void benchmarkExportResultsPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.benchmarkExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results3));
        }

        private void benchmarkExportResultsPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.benchmarkExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.benchmarkExportResultsPictureBox, "Export the results.");
        }

        private void benchmarkExportResultsPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.benchmarkExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.benchmarkExportResultsPictureBox, "Export the results.");
        }

        private void benchmarkExportResultsPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.benchmarkExportResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_export_results));
        }

        private void benchmarkClearResultsPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.benchmarkClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results3));
        }

        private void benchmarkClearResultsPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.benchmarkClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.benchmarkClearResultsPictureBox, "Clear the results.");
        }

        private void benchmarkClearResultsPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.benchmarkClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.benchmarkClearResultsPictureBox, "Clear the results.");
        }

        private void benchmarkClearResultsPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.benchmarkClearResultsPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear_results));
        }

        private void apiExportResultsPictureBox_Click(object sender, EventArgs e)
        {
            if (apiRichTextBox.Text == null || apiRichTextBox.Text == "")
            {
                MessageBox.Show("There are no results to export!", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            System.IO.Directory.CreateDirectory(Application.UserAppDataPath + @"\API_Readiness_Check");
            string path = Application.UserAppDataPath + @"\API_Readiness_Check\DataAnalysisTool_API_Check_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (TextWriter tw = new StreamWriter(fs))
                {
                    for (int i = 0; i < apiRichTextBox.Lines.Length; i++)
                    {
                        tw.WriteLine(apiRichTextBox.Lines[i]);
                    }
                    tw.WriteLine("EOF.");
                }
            }
            apiReadinessProgressBar.Value = 90;
            apiReadinessProgressBar.Value = 100;
            MessageBox.Show("API Readiness file has been created. \nLocation: " + path, "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            progressBar1.MarqueeAnimationSpeed = 0;
            Process.Start(path);
        }

        private void apiClearResultsPictureBox_Click(object sender, EventArgs e)
        {
            apiRichTextBox.Clear();
        }

        private void benchmarkExportResultsPictureBox_Click(object sender, EventArgs e)
        {
            if (benchmarkRichTextBox.Text == null || benchmarkRichTextBox.Text == "")
            {
                MessageBox.Show("There are no results to export!", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            System.IO.Directory.CreateDirectory(Application.UserAppDataPath + @"\Payout_Benchmarks");
            string path = Application.UserAppDataPath + @"\Payout_Benchmarks\DataAnalysisTool_PB_Data_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (TextWriter tw = new StreamWriter(fs))
                {
                    for (int i = 0; i < benchmarkRichTextBox.Lines.Length; i++)
                    {
                        tw.WriteLine(benchmarkRichTextBox.Lines[i]);
                    }
                    // setup for export
                    benchmarkDataGridView.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
                    benchmarkDataGridView.SelectAll();
                    // hiding row headers to avoid extra \t in exported text
                    var rowHeaders = benchmarkDataGridView.RowHeadersVisible;
                    benchmarkDataGridView.RowHeadersVisible = false;

                    // ! creating text from grid values
                    string content = benchmarkDataGridView.GetClipboardContent().GetText();

                    // restoring grid state
                    benchmarkDataGridView.ClearSelection();
                    benchmarkDataGridView.RowHeadersVisible = rowHeaders;
                    tw.WriteLine(content);
                    tw.WriteLine("EOF.");
                }
            }
            importFormatProgressBar.Value = 90;
            importFormatProgressBar.Value = 100;
            MessageBox.Show("Payout Benchmark file has been created. \nLocation: " + path, "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            progressBar1.MarqueeAnimationSpeed = 0;
            Process.Start(path);
        }

        private void benchmarkClearResultsPictureBox_Click(object sender, EventArgs e)
        {
            benchmarkRichTextBox.Clear();
        }

        private void sqlQueryGoPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.sqlQueryGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go3));
        }

        private void sqlQueryGoPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.sqlQueryGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.sqlQueryGoPictureBox, "Run the tool!");
        }

        private void sqlQueryGoPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.sqlQueryGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.sqlQueryGoPictureBox, "Run the tool!");
        }

        private void sqlQueryGoPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.sqlQueryGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
        }

        private void exportResultsPictureBox_Click(object sender, EventArgs e)
        {
            if (envChangesRichTextBox.Text == null || apiRichTextBox.Text == "")
            {
                MessageBox.Show("There are no results to export!", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            System.IO.Directory.CreateDirectory(Application.UserAppDataPath + @"\API_Readiness_Check");
            string path = Application.UserAppDataPath + @"\API_Readiness_Check\DataAnalysisTool_API_Check_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (TextWriter tw = new StreamWriter(fs))
                {
                    for (int i = 0; i < apiRichTextBox.Lines.Length; i++)
                    {
                        tw.WriteLine(apiRichTextBox.Lines[i]);
                    }
                    tw.WriteLine("EOF.");
                }
            }
            apiReadinessProgressBar.Value = 90;
            apiReadinessProgressBar.Value = 100;
            MessageBox.Show("Environment changes file has been created. \nLocation: " + path, "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            progressBar1.MarqueeAnimationSpeed = 0;
            Process.Start(path);
        }

        private void openInExcelPictureBox_Click(object sender, EventArgs e)
        {
            copyAlltoClipboard();
            Microsoft.Office.Interop.Excel.Application xlexcel;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;
            xlexcel = new Microsoft.Office.Interop.Excel.Application();
            xlexcel.Visible = true;
            xlWorkBook = xlexcel.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range CR = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, 1];
            CR.Select();
            xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
        }

        private void openInExcelPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.openInExcelPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_open_in_excel3));
        }

        private void openInExcelPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.openInExcelPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_open_in_excel2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.openInExcelPictureBox, "Open the table in Excel.");
        }

        private void openInExcelPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.openInExcelPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_open_in_excel));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.openInExcelPictureBox, "Open the table in Excel.");
        }

        private void openInExcelPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.openInExcelPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_open_in_excel));
        }

        private void legendButtonPictureBox_Click(object sender, EventArgs e)
        {
            DataGridViewLegend legend = new DataGridViewLegend();

            while (Application.OpenForms.Count > 1)
            {
                Application.OpenForms[Application.OpenForms.Count - 1].Close();
            }
            legend.ShowDialog();
        }

        private void legendButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.legendButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_table_legend3));
        }

        private void legendButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.legendButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_table_legend2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.legendButtonPictureBox, "Show the table legend.");
        }

        private void legendButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.legendButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_table_legend));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.legendButtonPictureBox, "Show the table legend.");
        }

        private void legendButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.legendButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_table_legend));
        }

        private void saveAsCsvButtonPictureBox_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            saveFileDialog1.Filter = "CSV|*.csv";
            if (this.saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                // create one file gridview.csv in writing mode using streamwriter
                StreamWriter sw = new StreamWriter("c:\\gridview.csv");
                // now add the gridview header in csv file suffix with "," delimeter except last one
                for (int i = 0; i < importedfileDataGridView.Columns.Count; i++)
                {
                    sw.Write(importedfileDataGridView.Columns[i].HeaderText);
                    if (i != importedfileDataGridView.Columns.Count)
                    {
                        sw.Write(",");
                    }
                }
                // add new line
                sw.Write(sw.NewLine);
                // iterate through all the rows within the gridview
                foreach (DataGridViewRow dr in importedfileDataGridView.Rows)
                {
                    // iterate through all colums of specific row
                    for (int i = 0; i < importedfileDataGridView.Columns.Count; i++)
                    {
                        // write particular cell to csv file
                        sw.Write(dr.Cells[i]);
                        if (i != importedfileDataGridView.Columns.Count)
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
            progressBar1.MarqueeAnimationSpeed = 0;
        }

        private void saveAsCsvButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.saveAsCsvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv_save3));
        }

        private void saveAsCsvButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.saveAsCsvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv_save2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.saveAsCsvButtonPictureBox, "Save as a CSV file.");
        }

        private void saveAsCsvButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.saveAsCsvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv_save));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.saveAsCsvButtonPictureBox, "Save as a CSV file.");
        }

        private void saveAsCsvButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.saveAsCsvButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_csv_save));
        }

        private void saveAsXmlButtonPictureBox_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            saveFileDialog1.Filter = "XML|*.xml";
            if (this.saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DataTable dt = (DataTable)this.importedfileDataGridView.DataSource;
                dt.WriteXml(this.saveFileDialog1.FileName, XmlWriteMode.WriteSchema);
            }
            progressBar1.MarqueeAnimationSpeed = 0;
        }

        private void saveAsXmlButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.saveAsXmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml_save3));
        }

        private void saveAsXmlButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.saveAsXmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml_save2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.saveAsXmlButtonPictureBox, "Save as an XML file.");
        }

        private void saveAsXmlButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.saveAsXmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml_save));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.saveAsXmlButtonPictureBox, "Save as an XML file.");
        }

        private void saveAsXmlButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.saveAsXmlButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_xml_save));
        }

        private void benchmarkGoPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.benchmarkGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go3));
        }

        private void benchmarkGoPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.benchmarkGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.benchmarkGoPictureBox, "Run the tool!");
        }

        private void benchmarkGoPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.benchmarkGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.benchmarkGoPictureBox, "Run the tool!");
        }

        private void benchmarkGoPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.benchmarkGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
        }

        private void selectAllCellLengthCheckerPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < cellLengthCheckerListBox.Items.Count; i++)
            {
                cellLengthCheckerListBox.SetSelected(i, true);
            }
        }

        private void clearAllCellLengthCheckerPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < cellLengthCheckerListBox.Items.Count; i++)
            {
                cellLengthCheckerListBox.SetSelected(i, false);
            }
        }

        private void cellLengthCheckerGoButtonPictureBox_Click(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
            if (checkToolsMaxLengthTextBox.Text.Length == 0)
            {
                MessageBox.Show("You did not enter a length!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            int length = int.Parse(checkToolsMaxLengthTextBox.Text);
            importFormatProgressBar.Value = 50;
            foreach (Object selecteditem in cellLengthCheckerListBox.SelectedItems)
            {
                a++;
                reqItem = selecteditem as String;
                int lengthCharacterCurIndex = cellLengthCheckerListBox.Items.IndexOf(reqItem);
                if (lengthCharacterCurIndex >= 0)
                {

                    for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                    {

                        var value = importedfileDataGridView.Rows[i].Cells[lengthCharacterCurIndex].Value.ToString();
                        //MessageBox.Show("value "+value+"reqitem "+reqItem);
                        if (value.Length > length)
                        {
                            importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[lengthCharacterCurIndex];
                            importFormatProgressBar.Value = 100;
                            MessageBox.Show("The value '" + value + "'" + " in column " + selecteditem + ", line " + (i + 1) + " is too long", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            return;
                        }
                    }
                }
            }
            if (a == 0)
            {
                MessageBox.Show("You did not select a column!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            importFormatProgressBar.Value = 100;
            MessageBox.Show("All columns/rows are under " + length, "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
        }

        private void clearAllNullCheckerButtonPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < nullCheckerListBox.Items.Count; i++)
            {
                nullCheckerListBox.SetSelected(i, false);
            }
        }

        private void selectAllNullCheckerButtonPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < nullCheckerListBox.Items.Count; i++)
            {
                nullCheckerListBox.SetSelected(i, true);
            }
        }

        private void nullCheckerGoButtonPictureBox_Click(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
            importFormatProgressBar.Value = 50;
            foreach (Object selecteditem in nullCheckerListBox.SelectedItems)
            {
                a++;
                reqItem = selecteditem as String;
                int nullCheckCurIndex = nullCheckerListBox.Items.IndexOf(reqItem);
                if (nullCheckCurIndex >= 0)
                {

                    for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                    {

                        var value = importedfileDataGridView.Rows[i].Cells[nullCheckCurIndex].Value.ToString();
                        if (string.IsNullOrWhiteSpace(value))
                        {
                            importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[nullCheckCurIndex];
                            importFormatProgressBar.Value = 100;
                            MessageBox.Show("NULL value found in column " + "'" + reqItem + "'" + " at line " + (i + 1), "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);

                            return;
                        }
                    }
                }
            }
            if (a == 0)
            {
                importFormatProgressBar.Value = 0;
                MessageBox.Show("You did not select a column!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            importFormatProgressBar.Value = 100;
            MessageBox.Show("no NULL value!", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }

        private void clearAllSpecialCharacterCheckerButtonPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < specialCharacterCheckerListBox.Items.Count; i++)
            {
                specialCharacterCheckerListBox.SetSelected(i, false);
            }
        }

        private void selectAllSpecialCharacterCheckerButtonPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < specialCharacterCheckerListBox.Items.Count; i++)
            {
                specialCharacterCheckerListBox.SetSelected(i, true);
            }
        }

        private void specialCharacterCheckerGoButtonPictureBox_Click(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
            String specialChar = checkToolsSpecialCharacterTextBox.Text;
            if (checkToolsSpecialCharacterTextBox.Text.Length == 0)
            {
                MessageBox.Show("You did not enter a special character!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            importFormatProgressBar.Value = 50;
            foreach (Object selecteditem in specialCharacterCheckerListBox.SelectedItems)
            {

                a++;
                reqItem = selecteditem as String;
                int specialCharacterCurIndex = specialCharacterCheckerListBox.Items.IndexOf(reqItem);
                if (specialCharacterCurIndex >= 0)
                {

                    for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                    {

                        var value = importedfileDataGridView.Rows[i].Cells[specialCharacterCurIndex].Value.ToString();
                        if (value.Contains(specialChar) == true)
                        {
                            importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[specialCharacterCurIndex];
                            importFormatProgressBar.Value = 100;
                            MessageBox.Show("'" + specialChar + "'" + " WAS found in the column " + "'" + selecteditem + "'" + " at line " + (i + 1), "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);

                            return;
                        }
                    }
                }
            }
            if (a == 0)
            {
                MessageBox.Show("You did not select a column!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                importFormatProgressBar.Value = 0;
                return;
            }
            importFormatProgressBar.Value = 100;
            MessageBox.Show("'" + specialChar + "'" + " WAS NOT FOUND!", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }

        private void clearAllDateCheckerButtonPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dateCheckerListBox.Items.Count; i++)
            {
                dateCheckerListBox.SetSelected(i, false);
            }
        }

        private void selectAllDateCheckerButtonPictureBox_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dateCheckerListBox.Items.Count; i++)
            {
                dateCheckerListBox.SetSelected(i, true);
            }
        }

        private void dateCheckerGoButtonPictureBox_Click(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
            importFormatProgressBar.Value = 50;
            foreach (Object selecteditem in dateCheckerListBox.SelectedItems)
            {
                a++;
                reqItem = selecteditem as String;
                int dateFormatCurIndex = dateCheckerListBox.Items.IndexOf(reqItem);
                if (dateFormatCurIndex >= 0)
                {
                    for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                    {
                        var value = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex].Value.ToString();
                        if (dateCheckerFindNullCheckbox.Checked)
                        {
                            if (value == " " || value == "" || value == null)
                            {
                                importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                importFormatProgressBar.Value = 100;
                                MessageBox.Show("NULL at line " + (i + 1) + "\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   NULL at line " + (i + 1) + "\r\nMake sure that the date is in the format: yyyymmdd");
                                return;
                            }
                        }

                        if (value.Length == 8)
                        {
                            try
                            {

                                int year = int.Parse(value.Substring(0, 4));
                                int month = int.Parse(value.Substring(4, 2));
                                int day = int.Parse(value.Substring(6, 2));

                                if (year > 2200)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    importFormatProgressBar.Value = 100;
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }

                                if (month > 12)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    importFormatProgressBar.Value = 100;
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }

                                if (month < 01)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    importFormatProgressBar.Value = 100;
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }

                                if (day > 31)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    importFormatProgressBar.Value = 100;
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }

                                if (day < 01)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    importFormatProgressBar.Value = 100;
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }
                            }
                            catch
                            {
                                importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                importFormatProgressBar.Value = 100;
                                MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The value has characters that are not numbers.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The value has characters that are not numbers.\r\nMake sure that the date is in the format: yyyymmdd");
                                return;
                            }
                        }
                        else
                        {
                            if (value.Length > 0)
                            {
                                importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                importFormatProgressBar.Value = 100;
                                MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyymmdd");
                                return;
                            }
                        }
                    }
                }
            }
            if (a == 0)
            {
                importFormatProgressBar.Value = 0;
                MessageBox.Show("You did not select a column!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            MessageBox.Show("Dates are OK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            importFormatProgressBar.Value = 100;
            systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Dates are OK");
            return;
        }

        private void cellLengthCheckerGoButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.cellLengthCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go3));
        }

        private void cellLengthCheckerGoButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.cellLengthCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.cellLengthCheckerGoButtonPictureBox, "Run the check.");
        }

        private void cellLengthCheckerGoButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.cellLengthCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.cellLengthCheckerGoButtonPictureBox, "Run the check.");
        }

        private void cellLengthCheckerGoButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.cellLengthCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
        }

        private void clearAllCellLengthCheckerPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.clearAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear3));
        }

        private void clearAllCellLengthCheckerPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.clearAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllCellLengthCheckerPictureBox, "Clear all.");
        }

        private void clearAllCellLengthCheckerPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.clearAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllCellLengthCheckerPictureBox, "Clear all.");
        }

        private void clearAllCellLengthCheckerPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.clearAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
        }

        private void selectAllCellLengthCheckerPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.selectAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all3));
        }

        private void selectAllCellLengthCheckerPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.selectAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllCellLengthCheckerPictureBox, "Select all.");
        }

        private void selectAllCellLengthCheckerPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.selectAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllCellLengthCheckerPictureBox, "Select all.");
        }

        private void selectAllCellLengthCheckerPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.selectAllCellLengthCheckerPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
        }

        private void nullCheckerGoButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.nullCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go3));
        }

        private void nullCheckerGoButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.nullCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.nullCheckerGoButtonPictureBox, "Run the check.");
        }

        private void nullCheckerGoButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.nullCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.nullCheckerGoButtonPictureBox, "Run the check.");
        }

        private void nullCheckerGoButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.nullCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
        }

        private void clearAllNullCheckerButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.clearAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear3));
        }

        private void clearAllNullCheckerButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.clearAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllNullCheckerButtonPictureBox, "Clear all.");
        }

        private void clearAllNullCheckerButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.clearAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllNullCheckerButtonPictureBox, "Clear all.");
        }

        private void clearAllNullCheckerButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.clearAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
        }

        private void selectAllNullCheckerButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.selectAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all3));
        }

        private void selectAllNullCheckerButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.selectAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllNullCheckerButtonPictureBox, "Select all.");
        }

        private void selectAllNullCheckerButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.selectAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllNullCheckerButtonPictureBox, "Select all.");
        }

        private void selectAllNullCheckerButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.selectAllNullCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
        }

        private void specialCharacterCheckerGoButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.specialCharacterCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go3));
        }

        private void specialCharacterCheckerGoButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.specialCharacterCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.specialCharacterCheckerGoButtonPictureBox, "Run the check.");
        }

        private void specialCharacterCheckerGoButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.specialCharacterCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.specialCharacterCheckerGoButtonPictureBox, "Run the check.");
        }

        private void specialCharacterCheckerGoButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.specialCharacterCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
        }

        private void clearAllSpecialCharacterCheckerButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.clearAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear3));
        }

        private void clearAllSpecialCharacterCheckerButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.clearAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllSpecialCharacterCheckerButtonPictureBox, "Clear all.");
        }

        private void clearAllSpecialCharacterCheckerButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.clearAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllSpecialCharacterCheckerButtonPictureBox, "Clear all.");
        }

        private void clearAllSpecialCharacterCheckerButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.clearAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
        }

        private void selectAllSpecialCharacterCheckerButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.selectAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all3));
        }

        private void selectAllSpecialCharacterCheckerButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.selectAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllSpecialCharacterCheckerButtonPictureBox, "Select all.");
        }

        private void selectAllSpecialCharacterCheckerButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.selectAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllSpecialCharacterCheckerButtonPictureBox, "Select all.");
        }

        private void selectAllSpecialCharacterCheckerButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.selectAllSpecialCharacterCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
        }

        private void dateCheckerGoButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.dateCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go3));
        }

        private void dateCheckerGoButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.dateCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.dateCheckerGoButtonPictureBox, "Run the check.");
        }

        private void dateCheckerGoButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.dateCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.dateCheckerGoButtonPictureBox, "Run the check.");
        }

        private void dateCheckerGoButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.dateCheckerGoButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources._2button_go));
        }

        private void clearAllDateCheckerButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.clearAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear3));
        }

        private void clearAllDateCheckerButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.clearAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllDateCheckerButtonPictureBox, "Clear all.");
        }

        private void clearAllDateCheckerButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.clearAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.clearAllDateCheckerButtonPictureBox, "Clear all.");
        }

        private void clearAllDateCheckerButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.clearAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_clear));
        }

        private void selectAllDateCheckerButtonPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.selectAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all3));
        }

        private void selectAllDateCheckerButtonPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.selectAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllDateCheckerButtonPictureBox, "Select all.");
        }

        private void selectAllDateCheckerButtonPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.selectAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.selectAllDateCheckerButtonPictureBox, "Select all.");
        }

        private void selectAllDateCheckerButtonPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.selectAllDateCheckerButtonPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_all));
        }

        private void fileSweepUploadFilesPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.fileSweepUploadFilesPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_upload_files3));
        }

        private void fileSweepUploadFilesPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.fileSweepUploadFilesPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_upload_files2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.fileSweepUploadFilesPictureBox, "Upload file(s).");
        }

        private void fileSweepUploadFilesPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.fileSweepUploadFilesPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_upload_files));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.fileSweepUploadFilesPictureBox, "Upload file(s).");
        }

        private void fileSweepUploadFilesPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.fileSweepUploadFilesPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_upload_files));
        }

        private void fileSweepGoPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.fileSweepGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go3));
        }

        private void fileSweepGoPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.fileSweepGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.fileSweepGoPictureBox, "Run the tool!");
        }

        private void fileSweepGoPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.fileSweepGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.fileSweepGoPictureBox, "Run the tool!");
        }

        private void fileSweepGoPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.fileSweepGoPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_go));
        }

        private void fileSweepUploadFilesPictureBox_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;

            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog() { ValidateNames = true, Multiselect = true })
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        fileSweepDataGridView.Columns.Add("FileName", "File Name");
                        DataGridViewColumn columnWidth = fileSweepDataGridView.Columns[0];
                        columnWidth.Width = 200;
                        foreach (String file in ofd.SafeFileNames)
                        {
                            fileSweepDataGridView.Rows.Add(file);
                        }
                        //ofd.FileNames gets the entire path and file
                        foreach (DataGridViewColumn column in fileSweepDataGridView.Columns)
                        {
                            column.SortMode = DataGridViewColumnSortMode.Automatic;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            var importedFileArray = importedfileDataGridView.Columns.Cast<DataGridViewColumn>()
                .Select(x => x.HeaderCell.Value.ToString().Trim()).ToArray();



            dateCheckerListBox.Items.Clear();
            specialCharacterCheckerListBox.Items.Clear();
            nullCheckerListBox.Items.Clear();
            cellLengthCheckerListBox.Items.Clear();
            int a = 0;
            for (int i = 0; i < importedFileArray.Length; i++)
            {
                a++;

                specialCharacterCheckerListBox.Items.Add(a + ". " + importedFileArray[i].ToString());
                dateCheckerListBox.Items.Add(a + ". " + importedFileArray[i].ToString());
                nullCheckerListBox.Items.Add(a + ". " + importedFileArray[i].ToString());
                cellLengthCheckerListBox.Items.Add(a + ". " + importedFileArray[i].ToString());
            }

            progressBar1.MarqueeAnimationSpeed = 0;
        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            DataGridView dgv = fileSweepDataGridView;
            try
            {
                int totalRows = dgv.Rows.Count;
                // get index of the row for the selected cell
                int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
                if (rowIndex == 0)
                    return;
                // get index of the column for the selected cell
                int colIndex = dgv.SelectedCells[0].OwningColumn.Index;
                DataGridViewRow selectedRow = dgv.Rows[rowIndex];
                dgv.Rows.Remove(selectedRow);
                dgv.Rows.Insert(rowIndex - 1, selectedRow);
                dgv.ClearSelection();
                dgv.Rows[rowIndex - 1].Cells[colIndex].Selected = true;
            }
            catch { }
        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            DataGridView dgv = fileSweepDataGridView;
            try
            {
                int totalRows = dgv.Rows.Count;
                // get index of the row for the selected cell
                int rowIndex = dgv.SelectedCells[0].OwningRow.Index;
                if (rowIndex == totalRows - 1)
                    return;
                // get index of the column for the selected cell
                int colIndex = dgv.SelectedCells[0].OwningColumn.Index;
                DataGridViewRow selectedRow = dgv.Rows[rowIndex];
                dgv.Rows.Remove(selectedRow);
                dgv.Rows.Insert(rowIndex + 1, selectedRow);
                dgv.ClearSelection();
                dgv.Rows[rowIndex + 1].Cells[colIndex].Selected = true;
            }
            catch { }
        }

        private void moveUpPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.moveUpPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_up3));
        }

        private void moveUpPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.moveUpPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_up2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.moveUpPictureBox, "Run the tool!");
        }

        private void moveUpPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.moveUpPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_up));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.moveUpPictureBox, "Run the tool!");
        }

        private void moveUpPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.moveUpPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_up));
        }

        private void moveDownPictureBox_MouseDown(object sender, MouseEventArgs e)
        {
            this.moveDownPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_down3));
        }

        private void moveDownPictureBox_MouseEnter(object sender, EventArgs e)
        {
            this.moveDownPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_down2));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.moveDownPictureBox, "Run the tool!");
        }

        private void moveDownPictureBox_MouseLeave(object sender, EventArgs e)
        {
            this.moveDownPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_down));
            System.Windows.Forms.ToolTip ToolTip = new System.Windows.Forms.ToolTip();
            ToolTip.SetToolTip(this.moveDownPictureBox, "Run the tool!");
        }

        private void moveDownPictureBox_MouseUp(object sender, MouseEventArgs e)
        {
            this.moveDownPictureBox.Image = ((System.Drawing.Image)(Properties.Resources.button_move_down));
        }
    }
}