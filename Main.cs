﻿using System;
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
using System.Linq;

namespace DataAnalysisTool
{

    public partial class DataAnalysisTool : Form
    {
        /*
         * ############################################################################################   
         * ############################################################################################
         * ####################PRODUCTION CODE BEGIN###################################################
         * ############################################################################################
         * ############################################################################################
        */

        //*********************************************************************************************
        //*********************************HEADER MENU*************************************************
        //*********************************************************************************************

        //------------------CC LOG OPEN START------------------------------------------------------
        private void cCDataToolLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            Process.Start(Application.UserAppDataPath + @"\Logs");
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar1.Refresh();
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

                        toolStripStatusLabel13.Text = ofd.FileName;
                        toolStripStatusLabel4.Text = importedfileDataGridView[0, importedfileDataGridView.Rows.Count - 1].Value.ToString();
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading XML: " + ofd.FileName + "...Done.");
                        toolStripStatusLabel2.Visible = true;
                        toolStripStatusLabel3.Visible = true;
                        toolStripStatusLabel4.Visible = true;
                        toolStripStatusLabel5.Visible = true;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar1.Refresh();
        }
        private void menu_Save_Xml_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            saveFileDialog1.Filter = "XML|*.xml";
            if (this.saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                DataTable dt = (DataTable)this.importedfileDataGridView.DataSource;
                dt.WriteXml(this.saveFileDialog1.FileName, XmlWriteMode.WriteSchema);
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar1.Refresh();
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
                    this.toolStripStatusLabel13.Text = openfile1.FileName;
                }
                {
                    string pathconn = "Provider = Microsoft.jet.OLEDB.4.0; Data source=" + toolStripStatusLabel13.Text + ";Extended Properties=\"Excel 8.0;HDR= yes;\";";
                    OleDbConnection conn = new OleDbConnection(pathconn);
                    OleDbDataAdapter MyDataAdapter = new OleDbDataAdapter("Select * from [Sheet1$]", conn);
                    DataTable dt = new DataTable();
                    MyDataAdapter.Fill(dt);
                    importedfileDataGridView.DataSource = dt;
                }
            }
            catch { }
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar1.Refresh();

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
            progressBar1.Refresh();
        }
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //Print the contents.
            e.Graphics.DrawImage(bitmap, 0, 0);
        }
        //------------------PRINT DOCUMENT END------------------------------------------------------

        //------------------CC LOGO CLICK START------------------------------------------------------
        private void ccLogo_Click1(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            System.Diagnostics.Process.Start("https://calliduscloud.com");
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar1.Refresh();
        }
        //------------------CC LOGO CLICK END------------------------------------------------------

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
                        toolStripStatusLabel13.Text = ofd.FileName;
                        toolStripStatusLabel13.Visible = true;
                        toolStripStatusLabel4.Text = importedfileDataGridView.Rows.Count.ToString();
                        toolStripStatusLabel2.Visible = true;
                        toolStripStatusLabel3.Visible = true;
                        toolStripStatusLabel4.Visible = true;
                        toolStripStatusLabel5.Visible = true;
                        toolStripStatusLabel12.Visible = true;
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading CSV: " + ofd.FileName + "...Done.");
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
            progressBar1.Refresh();
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
            progressBar1.Refresh();
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
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar1.Refresh();
        }
        //------------------EXIT APP ACTION END------------------------------------------------------


        //*********************************************************************************************
        //*********************************/HEADER MENU************************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************SHORTCUT TAB************************************************
        //*********************************************************************************************

        //------------------ACKTEKSOFT LOGIN START------------------------------------------------------
        private void acteksoft_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;

            acteksoft actek = new acteksoft();

            while (Application.OpenForms.Count > 1)
            {
                Application.OpenForms[Application.OpenForms.Count - 1].Close();
            }
            actek.ShowDialog();
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar1.Refresh();
        }
        //------------------ACKTEKSOFT LOGIN END------------------------------------------------------

        //*********************************************************************************************
        //*********************************/SHORTCUT TAB***********************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************CELL CHECK TAB**********************************************
        //*********************************************************************************************

        //------------------SELECT/CLEAR LIST BOX START------------------------------------------------------

        private void button17_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dateCheckerListBox.Items.Count; i++)
            {
                dateCheckerListBox.SetSelected(i, true);
            }
        }
        private void button18_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < specialCharacterCheckerListBox.Items.Count; i++)
            {
                specialCharacterCheckerListBox.SetSelected(i, true);
            }
        }
        private void button19_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < nullCheckerListBox.Items.Count; i++)
            {
                nullCheckerListBox.SetSelected(i, true);
            }
        }
        private void button20_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < cellLengthCheckerListBox.Items.Count; i++)
            {
                cellLengthCheckerListBox.SetSelected(i, true);
            }
        }
        private void button21_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dateCheckerListBox.Items.Count; i++)
            {
                dateCheckerListBox.SetSelected(i, false);
            }
        }
        private void button24_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < specialCharacterCheckerListBox.Items.Count; i++)
            {
                specialCharacterCheckerListBox.SetSelected(i, false);
            }
        }
        private void button23_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < nullCheckerListBox.Items.Count; i++)
            {
                nullCheckerListBox.SetSelected(i, false);
            }
        }
        private void button22_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < cellLengthCheckerListBox.Items.Count; i++)
            {
                cellLengthCheckerListBox.SetSelected(i, false);
            }
        }

        //------------------SELECT/CLEAR LIST BOX END------------------------------------------------------

        //------------------DATE CONVERTER START------------------------------------------------------
        private void dateConvert_Click1(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
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
                        if (checkBox1.Checked)
                        {
                            if (value == " " || value == "" || value == null)
                            {
                            importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                            MessageBox.Show("NULL at line " + (i + 1) + "\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   NULL at line " + (i + 1) + "\r\nMake sure that the date is in the format: yyyymmdd");
                                return;
                            }
                        }

                        if (value.Length == 8)
                        {
                            try {

                                int year = int.Parse(value.Substring(0, 4));
                                int month = int.Parse(value.Substring(4, 2));
                                int day = int.Parse(value.Substring(6, 2));

                                if (year > 2200)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }

                                if (month > 12)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }

                                if (month < 01)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }

                                if (day > 31)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }

                                if (day < 01)
                                {
                                    importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                    MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyymmdd");
                                    return;
                                }
                            }
                            catch {
                                importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The value has characters that are not numbers.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The value has characters that are not numbers.\r\nMake sure that the date is in the format: yyyymmdd");
                                return;
                            }
                            }
                        else
                        {
                            if (value.Length > 0)
                            {
                                importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[dateFormatCurIndex];
                                MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyymmdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyymmdd");
                                return;
                            }
                        }
                    }

                }
            }
            if (a == 0){
                MessageBox.Show("You did not select a column!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            MessageBox.Show("Dates are OK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Dates are OK");
            return;
        }
        //------------------DATE CONVERTER END------------------------------------------------------

        //------------------NULL CHECKER START------------------------------------------------------

        private void nullChecker_Click(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
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
                            MessageBox.Show("NULL value found in column " + "'" + reqItem + "'" + " at line " + (i + 1), "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);

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
            MessageBox.Show("no NULL value!", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }
        //------------------NULL CHECKER END------------------------------------------------------

        //------------------CELL LENGTH CHECKER START------------------------------------------------------

        private void cellLength_Click(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
            if (textBox4.Text.Length == 0)
            {
                MessageBox.Show("You did not enter a length!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
            int length = int.Parse(textBox4.Text);

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
            MessageBox.Show("All columns/rows are under " + length, "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);




        }

        //------------------CELL LENGTH CHECKER END------------------------------------------------------

        //------------------SPECIAL CHARACTER CHECKER START------------------------------------------------------

        private void specialCharacter_Click(object sender, EventArgs e)
        {
            int a = 0;
            String reqItem;
            String specialChar = textBox1.Text;
            if (textBox1.Text.Length == 0)
            {
                MessageBox.Show("You did not enter a special character!\r\nThe operation will now cancel.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                return;
            }
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
                        //MessageBox.Show("value "+value+"reqitem "+reqItem);
                        if (value.Contains(specialChar) == true)
                        {
                            importedfileDataGridView.CurrentCell = importedfileDataGridView.Rows[i].Cells[specialCharacterCurIndex];
                            MessageBox.Show("'" + specialChar + "'" + " WAS found in the column " + "'" + selecteditem + "'" + " at line " + (i + 1), "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);

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
            MessageBox.Show("'" + specialChar + "'" + " WAS NOT FOUND!", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }

        //------------------SPECIAL CHARACTER CHECKER END------------------------------------------------------

        //*********************************************************************************************
        //*********************************/CELL CHECK TAB*********************************************
        //*********************************************************************************************

        //*********************************************************************************************
        //*********************************GLOBAL******************************************************
        //*********************************************************************************************
        public DataAnalysisTool()
        {
            InitializeComponent();
            dateComboBox1.SelectedIndex = 12;
            dateComboBox2.SelectedIndex = 5;
            dateComboBox3.SelectedIndex = 1;
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            toolStripStatusLabel1.Text = @"TALLYCENTRAL\"+Environment.UserName;
        }

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

        //*********************************************************************************************
        //*********************************/GLOBAL*****************************************************
        //*********************************************************************************************


        private void toolStripStatusLabel4_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel4.Text = importedfileDataGridView.Rows.Count.ToString();
        }
        private void ssms_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            Ssms ssms = new Ssms();
            while (Application.OpenForms.Count > 1)
            {
                Application.OpenForms[Application.OpenForms.Count - 1].Close();
            }
            ssms.ShowDialog();
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar1.Refresh();
        }

        private void toolStripStatusLabel15_Click(object sender, EventArgs e)
        {
            DataGridViewLegend legend = new DataGridViewLegend();

            while (Application.OpenForms.Count > 1)
            {
                Application.OpenForms[Application.OpenForms.Count - 1].Close();
            }
            legend.ShowDialog();
        }



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
                        toolStripStatusLabel13.Text = ofd.FileName;
                        toolStripStatusLabel13.Visible = true;
                        toolStripStatusLabel4.Text = importedfileDataGridView.Rows.Count.ToString();
                        toolStripStatusLabel2.Visible = true;
                        toolStripStatusLabel3.Visible = true;
                        toolStripStatusLabel4.Visible = true;
                        toolStripStatusLabel5.Visible = true;
                        toolStripStatusLabel12.Visible = true;
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading TXT: " + ofd.FileName + "...Done.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar1.Refresh();
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
                        importedfileDataGridView.DataSource = ReadTxtPipe(ofd.FileName);
                        toolStripStatusLabel13.Text = ofd.FileName;
                        toolStripStatusLabel13.Visible = true;
                        toolStripStatusLabel4.Text = importedfileDataGridView.Rows.Count.ToString();
                        toolStripStatusLabel2.Visible = true;
                        toolStripStatusLabel3.Visible = true;
                        toolStripStatusLabel4.Visible = true;
                        toolStripStatusLabel5.Visible = true;
                        toolStripStatusLabel12.Visible = true;
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading TXT: " + ofd.FileName + "...Done.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar1.Refresh();
        }

        public DataTable ReadTxtPipe(string fileName)
        {
            DataTable dt = new DataTable();
            string[] columns = null;

            var lines = File.ReadAllLines(fileName);

            // assuming the first row contains the columns information
            if (lines.Count() > 0)
            {
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
            return dt;
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
        //------------------EXIT APP ACTION END------------------------------------------------------
        /*
         * ############################################################################################   
         * ############################################################################################
         * ####################PRODUCTION CODE END#####################################################
         * ############################################################################################
         * ############################################################################################
        */


        private void startAsyncButton_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.IsBusy != true)
            {
                // Start the asynchronous operation.
                backgroundWorker1.RunWorkerAsync();
            }
        }

        private void cancelAsyncButton_Click(object sender, EventArgs e)
        {
            if (backgroundWorker1.WorkerSupportsCancellation == true)
            {
                // Cancel the asynchronous operation.
                backgroundWorker1.CancelAsync();
            }
        }



        // This event handler is where the time-consuming work is done.
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            
            BackgroundWorker worker = sender as BackgroundWorker;

            for (int i = 1; i <= 10; i++)
            {
                if (worker.CancellationPending == true)
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    // Perform a time consuming operation and report progress.
                    System.Threading.Thread.Sleep(1);
                    worker.ReportProgress(i * 10);
                }
            }
        }


        // This event handler updates the progress.
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar2.Value = (e.ProgressPercentage);

        }

        // This event handler deals with the results of the background operation.
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //if (e.Cancelled == true)
            //{
            //    label1.Text = "Canceled!";
            //}
            //else if (e.Error != null)
            //{
            //    label1.Text = "Error: " + e.Error.Message;
            //}
            //else
            //{
            //    label1.Text = "Done!";
            //}
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.sap.com/index.html");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
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
            ToolTip2.SetToolTip(this.groupBox7, "Select your Server/Database/Import Format.");
        }

        private void reqListBox_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.reqListBox, "Select your required Import Format fields.");
        }

        private void groupBox1_MouseHover(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.groupBox1, "Select your required Import Format fields.");
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
            ToolTip2.SetToolTip(this.checkBox2, "Do you want to find NULLs in the date column?");
        }

        private void button6_MouseEnter(object sender, EventArgs e)
        {
            System.Windows.Forms.ToolTip ToolTip2 = new System.Windows.Forms.ToolTip();
            ToolTip2.SetToolTip(this.button6, "Run the tool!");
        }
    }
}