using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Collections;
using System.Configuration;
using System.Web;

namespace DataAnalysisTool
{
    public partial class DataAnalysisTool
    {

        //------------------MEDICARE CHECKER START------------------------------------------------------

        private void inProgramCheckToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ( databaseSelect.Text == "")

            {
                DialogResult result = MessageBox.Show("No database selected. \nThere will be no cross check with the database. Continue?", "Data Analysis Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                if (result == DialogResult.No)
                { return; }
            }

            if (databaseSelect.Text != "")

            {
                DialogResult result = MessageBox.Show("The DAT will check against the "+ifSelect.Text+" cross reference.\nContinue?", "Data Analysis Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                if (result == DialogResult.No)
                { return; }
                SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
                conn.Open();
                SqlCommand sc = new SqlCommand("use " + databaseSelect.Text + " select importformatid as name from ImportFormat", conn);
                SqlDataReader reader;
                try
                {

                    var select = "USE " + databaseSelect.Text + " select recval from CodSet where rectype='CMSPBP'";
                    var conn2 = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
                    var dataAdapter = new SqlDataAdapter(select, conn2);
                    var commandBuilder = new SqlCommandBuilder(dataAdapter);
                    var ds = new DataSet();
                    dataAdapter.Fill(ds);
                    stagedDataGridView.ReadOnly = true;
                    stagedDataGridView.DataSource = ds.Tables[0];
                    toolStripStatusLabel7.Text = stagedDataGridView.Rows.Count.ToString();
                    reader = sc.ExecuteReader();
                    DataTable dt = new DataTable();
                    toolStripStatusLabel10.Text = importformatDataGridView.Rows.Count.ToString();
                    conn.Close();

                    var array = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                             .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                    for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                    {
                        var value = importedfileDataGridView.Rows[i].Cells[2].Value.ToString();

                        if (array.Contains(value) == false)
                        {
                            MessageBox.Show("Error at line "+(i+1)+"."+"\n"+value + " from your imported file does not exist in the database.");
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line "+(i+1)+"."+"\n"+value + " from your imported file does not exist in the database.");
                            return;
                        }
                    }
                    richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading PBP for a cross check...Done.");
                }
                catch { return; }

                conn.Close();
            }


            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
            {
                    if (importedfileDataGridView.ColumnCount != 37)
                    {
                        MessageBox.Show("Medicare files need 37 columns. You have " + importedfileDataGridView.ColumnCount + ".", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return;
                    }
                    var value0 = importedfileDataGridView.Rows[i].Cells[0].Value.ToString();
                    var value1 = importedfileDataGridView.Rows[i].Cells[1].Value.ToString();
                    var value2 = importedfileDataGridView.Rows[i].Cells[2].Value.ToString();
                    var value3 = importedfileDataGridView.Rows[i].Cells[3].Value.ToString();
                    var value4 = importedfileDataGridView.Rows[i].Cells[4].Value.ToString();
                    var value5 = importedfileDataGridView.Rows[i].Cells[5].Value.ToString();
                    var value6 = importedfileDataGridView.Rows[i].Cells[6].Value.ToString();
                    var value7 = importedfileDataGridView.Rows[i].Cells[7].Value.ToString();
                    var value8 = importedfileDataGridView.Rows[i].Cells[8].Value.ToString();
                    var value9 = importedfileDataGridView.Rows[i].Cells[9].Value.ToString();
                    var value10 = importedfileDataGridView.Rows[i].Cells[10].Value.ToString();
                    var value11 = importedfileDataGridView.Rows[i].Cells[11].Value.ToString();
                    var value12 = importedfileDataGridView.Rows[i].Cells[12].Value.ToString();
                    var value13 = importedfileDataGridView.Rows[i].Cells[13].Value.ToString();
                    var value14 = importedfileDataGridView.Rows[i].Cells[14].Value.ToString();
                    var value15 = importedfileDataGridView.Rows[i].Cells[15].Value.ToString();
                    var value16 = importedfileDataGridView.Rows[i].Cells[16].Value.ToString();
                    var value17 = importedfileDataGridView.Rows[i].Cells[17].Value.ToString();
                    var value18 = importedfileDataGridView.Rows[i].Cells[18].Value.ToString();
                    var value19 = importedfileDataGridView.Rows[i].Cells[19].Value.ToString();
                    var value20 = importedfileDataGridView.Rows[i].Cells[20].Value.ToString();
                    var value21 = importedfileDataGridView.Rows[i].Cells[21].Value.ToString();
                    var value22 = importedfileDataGridView.Rows[i].Cells[22].Value.ToString();
                    var value23 = importedfileDataGridView.Rows[i].Cells[23].Value.ToString();
                    var value24 = importedfileDataGridView.Rows[i].Cells[24].Value.ToString();
                    var value25 = importedfileDataGridView.Rows[i].Cells[25].Value.ToString();
                    var value26 = importedfileDataGridView.Rows[i].Cells[26].Value.ToString();
                    var value27 = importedfileDataGridView.Rows[i].Cells[27].Value.ToString();
                    var value28 = importedfileDataGridView.Rows[i].Cells[28].Value.ToString();
                    var value29 = importedfileDataGridView.Rows[i].Cells[29].Value.ToString();
                    var value30 = importedfileDataGridView.Rows[i].Cells[30].Value.ToString();
                    var value31 = importedfileDataGridView.Rows[i].Cells[31].Value.ToString();
                    var value32 = importedfileDataGridView.Rows[i].Cells[32].Value.ToString();
                    var value33 = importedfileDataGridView.Rows[i].Cells[33].Value.ToString();
                    var value34 = importedfileDataGridView.Rows[i].Cells[34].Value.ToString();
                    var value35 = importedfileDataGridView.Rows[i].Cells[35].Value.ToString();
                    var value36 = importedfileDataGridView.Rows[i].Cells[36].Value.ToString();

                    //Required/Optional Check
                    if (string.IsNullOrWhiteSpace(value0))
                    {
                        MessageBox.Show("NULL value found in column #1 (CustomerId)  at line " + (i + 1) + " This is a required field.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   NULL value found in column #1 (CustomerId)  at line " + (i + 1) + " This is a required field.");
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(value1))
                    {
                        MessageBox.Show("NULL value found in column #2 (ContractNbr)  at line " + (i + 1) + " This is a required field.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   NULL value found in column #2 (ContractNbr)  at line " + (i + 1) + " This is a required field.");
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(value2))
                    {
                        MessageBox.Show("NULL value found in column #3 (PBP)  at line " + (i + 1) + " This is a required field.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   NULL value found in column #3 (PBP)  at line " + (i + 1) + " This is a required field.");
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(value3))
                    {
                        MessageBox.Show("NULL value found in column #4 (HICN)  at line " + (i + 1) + " This is a required field.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   NULL value found in column #4 (HICN)  at line " + (i + 1) + " This is a required field.");
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(value6))
                    {
                        MessageBox.Show("NULL value found in column #7 (DatEff)  at line " + (i + 1) + " This is a required field.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   NULL value found in column #7 (DatEff)  at line " + (i + 1) + " This is a required field.");
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(value8))
                    {
                        MessageBox.Show("NULL value found in column #9 (AppSignedDate)  at line " + (i + 1) + " This is a required field.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   NULL value found in column #9 (AppSignedDate)  at line " + (i + 1) + " This is a required field.");
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(value10))
                    {
                        MessageBox.Show("NULL value found in column #11 (Holder)  at line " + (i + 1) + " This is a required field.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   NULL value found in column #11 (Holder)  at line " + (i + 1) + " This is a required field.");
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(value23))
                    {
                        MessageBox.Show("NULL value found in column #24 (PolState)  at line " + (i + 1) + " This is a required field.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   NULL value found in column #24 (PolState)  at line " + (i + 1) + " This is a required field.");
                        return;
                    }
                    /////////////DATE PARSER/////////////
                    if (value6.Length == 8)
                    {
                        int year = int.Parse(value6.Substring(0, 4));
                        int month = int.Parse(value6.Substring(4, 2));
                        int day = int.Parse(value6.Substring(6, 2));

                        if (year > 2200)
                        {
                            MessageBox.Show("Error at column 7, line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month > 12)
                        {
                            MessageBox.Show("Error at column 7, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month < 01)
                        {
                            MessageBox.Show("Error at column 7, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day > 31)
                        {
                            MessageBox.Show("Error at column 7, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day < 01)
                        {
                            MessageBox.Show("Error at column 7, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }
                    }

                    if (value7.Length == 8)
                    {
                        int year = int.Parse(value7.Substring(0, 4));
                        int month = int.Parse(value7.Substring(4, 2));
                        int day = int.Parse(value7.Substring(6, 2));

                        if (year > 2200)
                        {
                            MessageBox.Show("Error at column 8, line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month > 12)
                        {
                            MessageBox.Show("Error at column 8, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month < 01)
                        {
                            MessageBox.Show("Error at column 8, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day > 31)
                        {
                            MessageBox.Show("Error at column 8, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day < 01)
                        {
                            MessageBox.Show("Error at column 8, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }
                    }
                    else if (value7.Length != 0)
                    {
                        MessageBox.Show("Error at column 8, line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd");
                        return;
                    }

                    if (value8.Length == 8)
                    {
                        int year = int.Parse(value8.Substring(0, 4));
                        int month = int.Parse(value8.Substring(4, 2));
                        int day = int.Parse(value8.Substring(6, 2));

                        if (year > 2200)
                        {
                            MessageBox.Show("Error at column 9, line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month > 12)
                        {
                            MessageBox.Show("Error at column 9, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month < 01)
                        {
                            MessageBox.Show("Error at column 9, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day > 31)
                        {
                            MessageBox.Show("Error at column 9, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day < 01)
                        {
                            MessageBox.Show("Error at column 9,  line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }
                    }
                    else if(value8.Length != 0)
                    {
                        MessageBox.Show("Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd");
                        return;
                    }

                    if (value9.Length == 8)
                    {
                        int year = int.Parse(value9.Substring(0, 4));
                        int month = int.Parse(value9.Substring(4, 2));
                        int day = int.Parse(value9.Substring(6, 2));

                        if (year > 2200)
                        {
                            MessageBox.Show("Error at column 10, line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month > 12)
                        {
                            MessageBox.Show("Error at column 10, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month < 01)
                        {
                            MessageBox.Show("Error at column 10, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day > 31)
                        {
                            MessageBox.Show("Error at column 10, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day < 01)
                        {
                            MessageBox.Show("Error at column 10, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }
                    }
                    else if (value9.Length != 0)
                    {
                        MessageBox.Show("Error at column 10, line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd");
                        return;
                    }

                    if (value21.Length == 8)
                    {
                        int year = int.Parse(value21.Substring(0, 4));
                        int month = int.Parse(value21.Substring(4, 2));
                        int day = int.Parse(value21.Substring(6, 2));

                        if (year > 2200)
                        {
                            MessageBox.Show("Error at column 22, line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month > 12)
                        {
                            MessageBox.Show("Error at column 22, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month < 01)
                        {
                            MessageBox.Show("Error at column 22, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day > 31)
                        {
                            MessageBox.Show("Error at column 22, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day < 01)
                        {
                            MessageBox.Show("Error at column 22, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }
                    }
                    else if (value21.Length != 0)
                    {
                        MessageBox.Show("Error at column 22, line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd");
                        return;
                    }

                    if (value36.Length == 8)
                    {
                        int year = int.Parse(value36.Substring(0, 4));
                        int month = int.Parse(value36.Substring(4, 2));
                        int day = int.Parse(value36.Substring(6, 2));

                        if (year > 2200)
                        {
                            MessageBox.Show("Error at column 37, line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month > 12)
                        {
                            MessageBox.Show("Error at column 37, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (month < 01)
                        {
                            MessageBox.Show("Error at column 37, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day > 31)
                        {
                            MessageBox.Show("Error at column 37, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }

                        if (day < 01)
                        {
                            MessageBox.Show("Error at column 37, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                            richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd");
                            return;
                        }
                    }
                    else if (value36.Length != 0)
                    {
                        MessageBox.Show("Error at column 37, line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Error at line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd");
                        return;
                    }


                    //Field Length Check
                    if (value0.Length > 30)
                    {
                        MessageBox.Show("column #1 (CustomerId)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #1 (CustomerId)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value1.Length > 10)
                    {
                        MessageBox.Show("column #2 (ContractNbr)  needs to be 10 or less characters.  At line " + (i + 1) + " you have a value that is " + value1.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #2 (ContractNbr)  needs to be 10 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value2.Length > 10)
                    {
                        MessageBox.Show("column #3 (PBP)  needs to be 10 or less characters.  At line " + (i + 1) + " you have a value that is " + value2.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #3 (PBP)  needs to be 10 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value3.Length > 20)
                    {
                        MessageBox.Show("column #4 (HICN)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value3.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #4 (HICN)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value4.Length > 30)
                    {
                        MessageBox.Show("column #5 (OED)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value4.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #5 (OED)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value5.Length > 30)
                    {
                        MessageBox.Show("column #6 (CMSOED)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value5.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #6 (CMSOED)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value6.Length > 30)
                    {
                        MessageBox.Show("column #7 (DatEff)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value6.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #7 (DatEff)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value7.Length > 30)
                    {
                        MessageBox.Show("column #8 (DatExp)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value7.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #8 (DatExp)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value8.Length > 30)
                    {
                        MessageBox.Show("column #9 (AppSignedDate)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value8.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #9 (AppSignedDate)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value9.Length > 30)
                    {
                        MessageBox.Show("column #10 (AppRcvDate)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value9.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #10 (AppRcvDate)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value10.Length > 60)
                    {
                        MessageBox.Show("column #11 (Holder)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value10.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #11 (Holder)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value11.Length > 40)
                    {
                        MessageBox.Show("column #12 (HolderFirstName)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value11.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #12 (HolderFirstName)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value12.Length > 16)
                    {
                        MessageBox.Show("column #13 (HolderMiddleInitial)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value12.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #13 (HolderMiddleInitial)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value13.Length > 60)
                    {
                        MessageBox.Show("column #14 (HolderLastName)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value13.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #14 (HolderLastName)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value14.Length > 60)
                    {
                        MessageBox.Show("column #15 (HolderStreet)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value14.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #15 (HolderStreet)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value15.Length > 30)
                    {
                        MessageBox.Show("column #16 (HolderStreet2)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value15.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #16 (HolderStreet2)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value16.Length > 40)
                    {
                        MessageBox.Show("column #17 (HolderCity)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value16.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #17 (HolderCity)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value17.Length > 6)
                    {
                        MessageBox.Show("column #18 (HolderState)  needs to be 6 or less characters.  At line " + (i + 1) + " you have a value that is " + value17.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #18 (HolderState)  needs to be 6 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value18.Length > 16)
                    {
                        MessageBox.Show("column #19 (HolderZip)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value18.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #19 (HolderZip)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value19.Length > 40)
                    {
                        MessageBox.Show("column #20 (CountyCode)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value19.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #20 (CountyCode)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value20.Length > 20)
                    {
                        MessageBox.Show("column #21 (HolderPhone)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value20.Length + " characters long.", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #21 (HolderPhone)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value21.Length > 30)
                    {
                        MessageBox.Show("column #22 (HolderDOB)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value21.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #22 (HolderDOB)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value22.Length > 20)
                    {
                        MessageBox.Show("column #23 (HolderSSN)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value22.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #23 (HolderSSN)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value23.Length > 30)
                    {
                        MessageBox.Show("column #24 (PolState)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value23.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #24 (PolState)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value24.Length > 8)
                    {
                        MessageBox.Show("column #25 (DualCoverage)  needs to be 8 or less characters.  At line " + (i + 1) + " you have a value that is " + value24.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #25 (DualCoverage)  needs to be 8 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value25.Length > 16)
                    {
                        MessageBox.Show("column #26 (BrokerId)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value25.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #26 (BrokerId)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value26.Length > 60)
                    {
                        MessageBox.Show("column #27 (TermType)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value26.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #27 (TermType)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value27.Length > 16)
                    {
                        MessageBox.Show("column #28 (ProCode)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value27.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #28 (ProCode)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value28.Length > 16)
                    {
                        MessageBox.Show("column #29 (BrokerId2)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value28.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #29 (BrokerId2)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value29.Length > 3.2)
                    {
                        MessageBox.Show("column #30 (PrimaryBrokerPct)  needs to be 3.2 or less characters.  At line " + (i + 1) + " you have a value that is " + value29.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #30 (PrimaryBrokerPct)  needs to be 3.2 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value30.Length > 3.2)
                    {
                        MessageBox.Show("column #31 (SecondaryBrokerPct)  needs to be 3.2 or less characters.  At line " + (i + 1) + " you have a value that is " + value30.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #31 (SecondaryBrokerPct)  needs to be 3.2 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value31.Length > 16)
                    {
                        MessageBox.Show("column #32 (ReferralId)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value31.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #32 (ReferralId)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value32.Length > 5)
                    {
                        MessageBox.Show("column #33 (BusType)  needs to be 5 or less characters.  At line " + (i + 1) + " you have a value that is " + value32.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #33 (BusType)  needs to be 5 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value33.Length > 30)
                    {
                        MessageBox.Show("column #34 (GroupId)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value33.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #34 (GroupId)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value34.Length > 40)
                    {
                        MessageBox.Show("column #35 (CustomerRegion)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value34.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #35 (CustomerRegion)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value35.Length > 20)
                    {
                        MessageBox.Show("column #36 (AppSource)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value35.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #36 (AppSource)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                    if (value36.Length > 30)
                    {
                        MessageBox.Show("column #37 (HolderDOD)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value36.Length + " characters long.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   column #37 (HolderDOD)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                        return;
                    }

                }
                {
                    MessageBox.Show("Medicare file is OK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Medicare file is OK");
                return;
                }
        }
        private void groupByColumnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (databaseSelect.Text == null || databaseSelect.Text=="" || databaseSelect.Text==" ")

            {
                DialogResult result =  MessageBox.Show("No database selected. \nThere will be no cross check with the database. Continue?", "Data Analysis Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                if (result ==DialogResult.No)
                { return; }
            }
            
            {
                System.IO.Directory.CreateDirectory(@"C:\Program Files (x86)\DataAnalysisTool\Medicare Error Files");
                string path = @"C:\Program Files (x86)\DataAnalysisTool\Medicare Error Files\DataAnalysisTool_MEF_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
                using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                {
                    using (TextWriter tw = new StreamWriter(fs))
                    {

                        tw.WriteLine("DataAnalysisTool - Beginning of Medicare Error File");
                        tw.WriteLine("Reading file...");
                        tw.WriteLine(".");
                        tw.WriteLine(".");
                        tw.WriteLine(".");
                        tw.WriteLine(".");

                        if (importedfileDataGridView.ColumnCount != 37)
                        {
                            tw.WriteLine("Medicare files need 37 columns. You have " + importedfileDataGridView.ColumnCount + ".");
                        }
                        //column 1 -required
                        try
                        {
                            tw.WriteLine("Column #1");

                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value0 = importedfileDataGridView.Rows[i].Cells[0].Value.ToString();

                                if (string.IsNullOrWhiteSpace(value0))
                                {
                                    tw.WriteLine("NULL value found in column #1 (CustomerId)  at line " + (i + 1) + ". This is a required field.");
                                }
                                if (value0.Length > 30)
                                {
                                    tw.WriteLine("column #1 (CustomerId)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value0.Length + " characters long.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #1 check...done."); }
                        //column 2 -required
                        try
                        {
                            tw.WriteLine("Column #2");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value1 = importedfileDataGridView.Rows[i].Cells[1].Value.ToString();

                                if (string.IsNullOrWhiteSpace(value1))
                                {
                                    tw.WriteLine("NULL value found in column #2 (ContractNbr)  at line " + (i + 1) + ". This is a required field.");
                                }
                                if (value1.Length > 10)
                                {
                                    tw.WriteLine("column #2 (ContractNbr)  needs to be 10 or less characters.  At line " + (i + 1) + " you have a value that is " + value1.Length + " characters long.");
                                    return;
                                }
                            }
                        }
                        catch { tw.WriteLine("column #2 check...done."); }
                        //column 3 -required
                        try
                        {
                            tw.WriteLine("Column #3");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value2 = importedfileDataGridView.Rows[i].Cells[2].Value.ToString();
                                if (string.IsNullOrWhiteSpace(value2))
                                {
                                    tw.WriteLine("NULL value found in column #3 (PBP)  at line " + (i + 1) + ". This is a required field.");
                                }
                                if (value2.Length > 10)
                                {
                                    tw.WriteLine("column #3 (PBP)  needs to be 10 or less characters.  At line " + (i + 1) + " you have a value that is " + value2.Length + " characters long.");
                                    return;
                                }
                            }
                        }
                        catch { tw.WriteLine("column #3 check...done."); }
                        //column 4 -required
                        try
                        {
                            tw.WriteLine("Column #4");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value3 = importedfileDataGridView.Rows[i].Cells[3].Value.ToString();
                                if (string.IsNullOrWhiteSpace(value3))
                                {
                                    tw.WriteLine("NULL value found in column #4 (HICN)  at line " + (i + 1) + ". This is a required field.");
                                }
                                if (value3.Length > 20)
                                {
                                    tw.WriteLine("column #4 (HICN)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value3.Length + " characters long.");
                                    return;
                                }
                            }
                        }
                        catch { tw.WriteLine("column #4 check...done."); }
                        //column 5 -not required
                        try
                        {
                            tw.WriteLine("Column #5");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value4 = importedfileDataGridView.Rows[i].Cells[4].Value.ToString();

                                if (value4.Length > 20)
                                {
                                    tw.WriteLine("column #5 (OED)  int");
                                    return;
                                }

                                if (value4.Length == 8)
                                {
                                    int year = int.Parse(value4.Substring(0, 4));
                                    int month = int.Parse(value4.Substring(4, 2));
                                    int day = int.Parse(value4.Substring(6, 2));

                                    if (year > 2200)
                                    {
                                        tw.WriteLine("Error at column 5, line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month > 12)
                                    {
                                        tw.WriteLine("Error at column 5, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month < 01)
                                    {
                                        tw.WriteLine("Error at column 5, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day > 31)
                                    {
                                        tw.WriteLine("Error at column 5, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day < 01)
                                    {
                                        tw.WriteLine("Error at column 5, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }
                                }
                                else if (value4.Length != 0)
                                {
                                    tw.WriteLine("Error at column 5, line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                }
                            }
                        }
                        catch { tw.WriteLine("column #5 check...done."); }
                        //column 6 -not required
                        try
                        {
                            tw.WriteLine("Column #6");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value5 = importedfileDataGridView.Rows[i].Cells[5].Value.ToString();

                                if (value5.Length > 20)
                                {
                                    tw.WriteLine("column #6 (CMSOED)  int");
                                }

                                if (value5.Length == 8)
                                {
                                    int year = int.Parse(value5.Substring(0, 4));
                                    int month = int.Parse(value5.Substring(4, 2));
                                    int day = int.Parse(value5.Substring(6, 2));

                                    if (year > 2200)
                                    {
                                        tw.WriteLine("Error at column 6, line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month > 12)
                                    {
                                        tw.WriteLine("Error at column 6, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month < 01)
                                    {
                                        tw.WriteLine("Error at column 6, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day > 31)
                                    {
                                        tw.WriteLine("Error at column 6, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day < 01)
                                    {
                                        tw.WriteLine("Error at column 6, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }
                                }
                                else if (value5.Length != 0)
                                {
                                    tw.WriteLine("Error at column 6, line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                }
                            }
                        }
                        catch { tw.WriteLine("column #6 check...done."); }
                        //column 7 -required
                        try
                        {
                            tw.WriteLine("Column #7");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value6 = importedfileDataGridView.Rows[i].Cells[6].Value.ToString();

                                if (string.IsNullOrWhiteSpace(value6))
                                {

                                    tw.WriteLine("NULL value found in column #7 (DatEff)  at line " + (i + 1) + ". This is a required field.");
                                }
                                if (value6.Length > 20)
                                {
                                    tw.WriteLine("column #7 (DatEff)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value6.Length + " characters long.");
                                }

                                if (value6.Length == 8)
                                {
                                    int year = int.Parse(value6.Substring(0, 4));
                                    int month = int.Parse(value6.Substring(4, 2));
                                    int day = int.Parse(value6.Substring(6, 2));

                                    if (year > 2200)
                                    {
                                        tw.WriteLine("Error at column 7, line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month > 12)
                                    {
                                        tw.WriteLine("Error at column 7, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month < 01)
                                    {
                                        tw.WriteLine("Error at column 7, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day > 31)
                                    {
                                        tw.WriteLine("Error at column 7, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day < 01)
                                    {
                                        tw.WriteLine("Error at column 7, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }
                                }
                                else if (value6.Length != 0)
                                {
                                    tw.WriteLine("Error at column 7, line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                }
                            }
                        }
                        catch { tw.WriteLine("column #7 check...done."); }
                        //column 8 -not required
                        try
                        {
                            tw.WriteLine("Column #8");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value7 = importedfileDataGridView.Rows[i].Cells[7].Value.ToString();

                                if (value7.Length > 20)
                                {
                                    tw.WriteLine("column #8 (DatExp)  int");
                                }

                                if (value7.Length == 8)
                                {
                                    int year = int.Parse(value7.Substring(0, 4));
                                    int month = int.Parse(value7.Substring(4, 2));
                                    int day = int.Parse(value7.Substring(6, 2));

                                    if (year > 2200)
                                    {
                                        tw.WriteLine("Error at column 8, line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month > 12)
                                    {
                                        tw.WriteLine("Error at column 8, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month < 01)
                                    {
                                        tw.WriteLine("Error at column 8, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day > 31)
                                    {
                                        tw.WriteLine("Error at column 8, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day < 01)
                                    {
                                        tw.WriteLine("Error at column 8, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }
                                }
                                else if (value7.Length != 0)
                                {
                                    tw.WriteLine("Error at column 8, line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                }
                            }
                        }
                        catch { tw.WriteLine("column #8 check...done."); }
                        //column 9 -required
                        try
                        {
                            tw.WriteLine("Column #9");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value8 = importedfileDataGridView.Rows[i].Cells[8].Value.ToString();

                                if (string.IsNullOrWhiteSpace(value8))
                                {
                                    tw.WriteLine("NULL value found in column #9 (AppSignedDate)  at line " + (i + 1) + ". This is a required field.");

                                }
                                if (value8.Length > 20)
                                {
                                    tw.WriteLine("column #9 (AppSignedDate)  int-length for this?");
                                }

                                if (value8.Length == 8)
                                {
                                    int year = int.Parse(value8.Substring(0, 4));
                                    int month = int.Parse(value8.Substring(4, 2));
                                    int day = int.Parse(value8.Substring(6, 2));

                                    if (year > 2200)
                                    {
                                        tw.WriteLine("Error at column 9, line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month > 12)
                                    {
                                        tw.WriteLine("Error at column 9, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month < 01)
                                    {
                                        tw.WriteLine("Error at column 9, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day > 31)
                                    {
                                        tw.WriteLine("Error at column 9, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day < 01)
                                    {
                                        tw.WriteLine("Error at column 9, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }
                                }
                                else if (value8.Length != 0)
                                {
                                    tw.WriteLine("Error at column 9, line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                }
                            }
                        }
                        catch { tw.WriteLine("column #9 check...done."); }
                        //column 10 -not required
                        try
                        {
                            tw.WriteLine("Column #10");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value9 = importedfileDataGridView.Rows[i].Cells[9].Value.ToString();

                                if (value9.Length > 20)
                                {
                                    tw.WriteLine("column #10 (AppRcvDate)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value9.Length + " characters long.");
                                }

                                if (value9.Length == 8)
                                {
                                    int year = int.Parse(value9.Substring(0, 4));
                                    int month = int.Parse(value9.Substring(4, 2));
                                    int day = int.Parse(value9.Substring(6, 2));

                                    if (year > 2200)
                                    {
                                        tw.WriteLine("Error at column 10, line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month > 12)
                                    {
                                        tw.WriteLine("Error at column 10, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month < 01)
                                    {
                                        tw.WriteLine("Error at column 10, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day > 31)
                                    {
                                        tw.WriteLine("Error at column 10, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day < 01)
                                    {
                                        tw.WriteLine("Error at column 10, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }
                                }
                                else if (value9.Length != 0)
                                {
                                    tw.WriteLine("Error at column 10, line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                }
                            }
                        }
                        catch { tw.WriteLine("column #10 check...done."); }
                        //column 11 -required
                        try
                        {
                            tw.WriteLine("Column #11");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value10 = importedfileDataGridView.Rows[i].Cells[10].Value.ToString();

                                if (string.IsNullOrWhiteSpace(value10))
                                {
                                    tw.WriteLine("NULL value found in column #11 (Holder)  at line " + (i + 1) + ". This is a required field.");
                                }
                                if (value10.Length > 60)
                                {
                                    tw.WriteLine("column #11 (Holder)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value10.Length + " characters long.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #11 check...done."); }
                        //column 12 -not required
                        try
                        {
                            tw.WriteLine("Column #12");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value11 = importedfileDataGridView.Rows[i].Cells[11].Value.ToString();

                                if (value11.Length > 60)
                                {
                                    tw.WriteLine("column #12 (HolderFirstName)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value11.Length + " characters long.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #12 check...done."); }
                        //column 13 -not required
                        try
                        {
                            tw.WriteLine("Column #13");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value12 = importedfileDataGridView.Rows[i].Cells[12].Value.ToString();

                                if (value12.Length > 60)
                                {
                                    tw.WriteLine("column #13 (HolderMiddleInitial)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value12.Length + " characters long.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #13 check...done."); }
                        //column 14 -not required
                        try
                        {
                            tw.WriteLine("Column #14");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value13 = importedfileDataGridView.Rows[i].Cells[13].Value.ToString();

                                if (value13.Length > 60)
                                {
                                    tw.WriteLine("column #14 (HolderLastName)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value13.Length + " characters long.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #14 check...done."); }
                        //column 15 -not required
                        try
                        {
                            tw.WriteLine("Column #15");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value14 = importedfileDataGridView.Rows[i].Cells[14].Value.ToString();

                                if (value14.Length > 60)
                                {
                                    tw.WriteLine("column #15 (HolderStreet)  needs to be 60 or less characters.  At line " + (i + 1) + " you have a value that is " + value14.Length + " characters long.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #15 check...done."); }
                        //column 16 -not required
                        try
                        {
                            tw.WriteLine("Column #16");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value15 = importedfileDataGridView.Rows[i].Cells[15].Value.ToString();

                                if (value15.Length > 60)
                                {
                                    tw.WriteLine("column #16 (HolderStreet2)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value15.Length + " characters long.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #16 check...done."); }
                        //column 17 -not required
                        try
                        {
                            tw.WriteLine("Column #17");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value16 = importedfileDataGridView.Rows[i].Cells[16].Value.ToString();

                                if (value16.Length > 60)
                                {
                                    tw.WriteLine("column #17 (HolderCity)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value16.Length + " characters long.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #17 check...done."); }
                        //column 18 -not required
                        try
                        {
                            tw.WriteLine("Column #18");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value17 = importedfileDataGridView.Rows[i].Cells[17].Value.ToString();

                                if (value17.Length > 60)
                                {
                                    tw.WriteLine("column #18 (HolderState)  needs to be 6 or less characters.  At line " + (i + 1) + " you have a value that is " + value17.Length + " characters long.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #18 check...done."); }
                        //column 19 -not required
                        try
                        {
                            tw.WriteLine("Column #19");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value18 = importedfileDataGridView.Rows[i].Cells[18].Value.ToString();

                                if (value18.Length > 60)
                                {
                                    tw.WriteLine("column #19 (HolderZip)  needs to be 16 or less characters.  At line " + (i + 1) + " you have a value that is " + value18.Length + " characters long.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #19 check...done."); }
                        //column 20 -not required
                        try
                        {
                            tw.WriteLine("Column #20");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value19 = importedfileDataGridView.Rows[i].Cells[19].Value.ToString();

                                if (value19.Length > 60)
                                {
                                    tw.WriteLine("column #20 (CountyCode)  needs to be 40 or less characters.  At line " + (i + 1) + " you have a value that is " + value19.Length + " characters long.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #20 check...done."); }
                        //column 21 -not required
                        try
                        {
                            tw.WriteLine("Column #21");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value20 = importedfileDataGridView.Rows[i].Cells[20].Value.ToString();

                                if (value20.Length > 60)
                                {
                                    tw.WriteLine("column #21 (HolderPhone)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value20.Length + " characters long.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #21 check...done."); }
                        //column 22 -not required
                        try
                        {
                            tw.WriteLine("Column #22");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value21 = importedfileDataGridView.Rows[i].Cells[21].Value.ToString();

                                if (value21.Length > 60)
                                {
                                    tw.WriteLine("column #22 (HolderDOB)  needs to be 30 or less characters.  At line " + (i + 1) + " you have a value that is " + value21.Length + " characters long.");
                                }

                                if (value21.Length == 8)
                                {
                                    int year = int.Parse(value21.Substring(0, 4));
                                    int month = int.Parse(value21.Substring(4, 2));
                                    int day = int.Parse(value21.Substring(6, 2));

                                    if (year > 2200)
                                    {
                                        tw.WriteLine("Error at column 22, line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month > 12)
                                    {
                                        tw.WriteLine("Error at column 22, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month < 01)
                                    {
                                        tw.WriteLine("Error at column 22, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day > 31)
                                    {
                                        tw.WriteLine("Error at column 22, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day < 01)
                                    {
                                        tw.WriteLine("Error at column 22, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }
                                }
                                else if (value21.Length != 0)
                                {
                                    tw.WriteLine("Error at column 8, line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                }
                            }
                        }
                        catch { tw.WriteLine("column #22 check...done."); }
                        //column 23 -not required
                        try
                        {
                            tw.WriteLine("Column #23");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value22 = importedfileDataGridView.Rows[i].Cells[22].Value.ToString();

                                if (value22.Length > 60)
                                {
                                    tw.WriteLine("column #23 (HolderSSN)  needs to be 20 or less characters.  At line " + (i + 1) + " you have a value that is " + value22.Length + " characters long.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #23 check...done."); }
                        //column 24 -required
                        try
                        {
                            tw.WriteLine("Column #24");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value23 = importedfileDataGridView.Rows[i].Cells[23].Value.ToString();

                                if (string.IsNullOrWhiteSpace(value23))
                                {
                                    tw.WriteLine("NULL value found in column #24 (PolState)  at line " + (i + 1) + ". This is a required field.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #24 check...done."); }
                        //column 25 -not required
                        try
                        {
                            tw.WriteLine("Column #25");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value24 = importedfileDataGridView.Rows[i].Cells[24].Value.ToString();

                                if (value24.Length > 60)
                                {
                                    tw.WriteLine("NULL value found in column #25 (DualCoverage)  at line " + (i + 1) + ". This is a required field.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #25 check...done."); }
                        //column 26 -not required
                        try
                        {
                            tw.WriteLine("Column #26");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value25 = importedfileDataGridView.Rows[i].Cells[25].Value.ToString();

                                if (value25.Length > 60)
                                {
                                    tw.WriteLine("NULL value found in column #26 (BrokerId)  at line " + (i + 1) + ". This is a required field.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #26 check...done."); }
                        //column 27 -not required
                        try
                        {
                            tw.WriteLine("Column #27");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value26 = importedfileDataGridView.Rows[i].Cells[26].Value.ToString();

                                if (value26.Length > 60)
                                {
                                    tw.WriteLine("NULL value found in column #27 (TermType)  at line " + (i + 1) + ". This is a required field.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #27 check...done."); }
                        //column 28 -not required
                        try
                        {
                            tw.WriteLine("Column #28");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value27 = importedfileDataGridView.Rows[i].Cells[27].Value.ToString();

                                if (value27.Length > 60)
                                {
                                    tw.WriteLine("NULL value found in column #28 (ProCode)  at line " + (i + 1) + ". This is a required field.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #28 check...done."); }
                        //column 29 -not required
                        try
                        {
                            tw.WriteLine("Column #29");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value28 = importedfileDataGridView.Rows[i].Cells[28].Value.ToString();

                                if (value28.Length > 60)
                                {
                                    tw.WriteLine("NULL value found in column #29 (BrokerId2)  at line " + (i + 1) + ". This is a required field.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #29 check...done."); }
                        //column 30 -not required
                        try
                        {
                            tw.WriteLine("Column #30");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value29 = importedfileDataGridView.Rows[i].Cells[29].Value.ToString();

                                if (value29.Length > 60)
                                {
                                    tw.WriteLine("NULL value found in column #30 (PrimaryBrokerPct)  at line " + (i + 1) + ". This is a required field.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #30 check...done."); }
                        //column 31 -not required
                        try
                        {
                            tw.WriteLine("Column #31");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value30 = importedfileDataGridView.Rows[i].Cells[30].Value.ToString();

                                if (value30.Length > 60)
                                {
                                    tw.WriteLine("NULL value found in column #31 (SecondaryBrokerPct)  at line " + (i + 1) + ". This is a required field.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #31 check...done."); }
                        //column 32 -not required
                        try
                        {
                            tw.WriteLine("Column #32");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value31 = importedfileDataGridView.Rows[i].Cells[31].Value.ToString();

                                if (value31.Length > 60)
                                {
                                    tw.WriteLine("NULL value found in column #32 (ReferralId)  at line " + (i + 1) + ". This is a required field.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #32 check...done."); }
                        //column 33 -not required
                        try
                        {
                            tw.WriteLine("Column #33");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value32 = importedfileDataGridView.Rows[i].Cells[32].Value.ToString();

                                if (value32.Length > 60)
                                {
                                    tw.WriteLine("NULL value found in column #33 (BusType)  at line " + (i + 1) + ". This is a required field.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #33 check...done."); }
                        //column 34 -not required
                        try
                        {
                            tw.WriteLine("Column #34");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value33 = importedfileDataGridView.Rows[i].Cells[33].Value.ToString();

                                if (value33.Length > 60)
                                {
                                    tw.WriteLine("NULL value found in column #34 (GroupId)  at line " + (i + 1) + ". This is a required field.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #34 check...done."); }
                        //column 35 -not required
                        try
                        {
                            tw.WriteLine("Column #35");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value34 = importedfileDataGridView.Rows[i].Cells[34].Value.ToString();

                                if (value34.Length > 60)
                                {
                                    tw.WriteLine("NULL value found in column #35 (CustomerRegion)  at line " + (i + 1) + ". This is a required field.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #35 check...done."); }
                        //column 36 -not required
                        try
                        {
                            tw.WriteLine("Column #36");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value35 = importedfileDataGridView.Rows[i].Cells[35].Value.ToString();

                                if (value35.Length > 60)
                                {
                                    tw.WriteLine("NULL value found in column #36 (AppSource)  at line " + (i + 1) + ". This is a required field.");
                                }
                            }
                        }
                        catch { tw.WriteLine("column #36 check...done."); }
                        //column 37 -not required
                        try
                        {
                            tw.WriteLine("Column #37");
                            for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                            {
                                var value36 = importedfileDataGridView.Rows[i].Cells[36].Value.ToString();

                                if (value36.Length > 60)
                                {
                                    tw.WriteLine("NULL value found in column #37 (HolderDOD)  at line " + (i + 1) + ". This is a required field.");
                                }

                                if (value36.Length == 8)
                                {
                                    int year = int.Parse(value36.Substring(0, 4));
                                    int month = int.Parse(value36.Substring(4, 2));
                                    int day = int.Parse(value36.Substring(6, 2));

                                    if (year > 2200)
                                    {
                                        tw.WriteLine("Error at column 37, line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month > 12)
                                    {
                                        tw.WriteLine("Error at column 37, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (month < 01)
                                    {
                                        tw.WriteLine("Error at column 37, line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day > 31)
                                    {
                                        tw.WriteLine("Error at column 37, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }

                                    if (day < 01)
                                    {
                                        tw.WriteLine("Error at column 37, line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                    }
                                }
                                else if (value36.Length != 0)
                                {
                                    tw.WriteLine("Error at column 37, line " + (i + 1) + "\r\n" + "The year is not 8 digits.\r\nMake sure that the date is in the format: yyyyMMdd", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                }
                            }
                        }
                        catch { tw.WriteLine("column #37 check...done."); }
                        tw.WriteLine("EOF.");
                    }
                }
                MessageBox.Show(@"Medicare error file has been created. \nLocation: C:\Program Files (x86)\DataAnalysisTool\Medicare Error Files", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + @">>>   Medicare error file has been created. Location: C:\Program Files (x86)\DataAnalysisTool\Medicare Error Files");
            }
            //------------------MEDICARE CHECKER END------------------------------------------------------

        }
        private void groupByErrorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //global vars
            progressBar1.MarqueeAnimationSpeed = 1;
            var ifCount = "USE " + databaseSelect.Text + " SELECT IMFF.FieldSeq FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null order by imff.FieldSeq";
            

            if (importedfileDataGridView.Rows.Count == 0)

            {
               MessageBox.Show("No file imported. \nPlease open a file.", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                progressBar1.Refresh();
                return; 
            }

            if (databaseSelect.Text == "")

            {
                DialogResult result = MessageBox.Show("No database selected. \nThere will be no cross check with the database. Continue?", "Data Analysis Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                if (result == DialogResult.No)
                {
                    progressBar1.MarqueeAnimationSpeed = 0;
                    progressBar1.Refresh();
                    return;
                }
            }

            if (databaseSelect.Text != "")
            {

                DialogResult result2 = MessageBox.Show("The DAT will check against the " + ifSelect.Text + " Import Format.\nContinue?", "Data Analysis Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
                if (result2 == DialogResult.No)
                {
                    progressBar1.MarqueeAnimationSpeed = 0;
                    progressBar1.Refresh();
                    return;
                }
            }

            {
                System.IO.Directory.CreateDirectory(@"C:\Program Files (x86)\DataAnalysisTool\Import Format Error Files");
                string path = @"C:\Program Files (x86)\DataAnalysisTool\Import Format Error Files\DataAnalysisTool_IFEF_" + DateTime.Now.ToString("MM_dd_yyyy_HHmmss") + ".txt";
                using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                {
                    using (TextWriter tw = new StreamWriter(fs))
                    {
                        tw.WriteLine("DataAnalysisTool - Beginning of Import Format Error File");
                        tw.WriteLine("Reading file...");
                        tw.WriteLine(".");
                        tw.WriteLine(".");
                        tw.WriteLine("Server: " + serverSelect.Text);
                        tw.WriteLine("Database: "+databaseSelect.Text);



                        if (databaseSelect.Text != "")
                        {
                            if (importedfileDataGridView.ColumnCount != importformatDataGridView.RowCount)
                            {
                                tw.WriteLine("This Import Format requires " + importformatDataGridView.RowCount + " columns. You have " + importedfileDataGridView.ColumnCount + ".");
                                tw.WriteLine("This operation has ended. Please correct the column count issue.");
                                MessageBox.Show("Import Format error file has been created. \nLocation: C:\\Program Files (x86)\\DataAnalysisTool\\Medicare Error Files", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + @">>>   Import Format error file has been created. Location: C:\Program Files (x86)\DataAnalysisTool\Medicare Error Files");
                                progressBar1.MarqueeAnimationSpeed = 0;
                                progressBar1.Refresh();
                                return;
                            }

                            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
                            conn.Open();
                            SqlCommand sc = new SqlCommand("use " + databaseSelect.Text + " select importformatid as name from ImportFormat", conn);
                            try
                            {
                                var selectCodeType = "USE " + databaseSelect.Text + " SELECT ef.codetype FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null and ef.valuetype=1 order by imff.FieldSeq";
                                var dataAdapter = new SqlDataAdapter(selectCodeType, conn);
                                var ds = new DataSet();
                                dataAdapter.Fill(ds);
                                stagedDataGridView.DataSource = ds.Tables[0];
                                var codeArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

                                var selectFieldSeq = "USE " + databaseSelect.Text + " SELECT IMFF.FieldSeq FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null and ef.valuetype=1 order by imff.FieldSeq";
                                var dataAdapter3 = new SqlDataAdapter(selectFieldSeq, conn);
                                var ds3 = new DataSet();
                                dataAdapter3.Fill(ds3);
                                stagedDataGridView.DataSource = ds3.Tables[0];
                                var fieldsThatAreCodesArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

                                var selectMaxLength = "USE " + databaseSelect.Text + " SELECT ef.FldName FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null and ef.MaxLength is not null order by imff.FieldSeq";
                                var dataAdapter4 = new SqlDataAdapter(selectMaxLength, conn);
                                var ds4 = new DataSet();
                                dataAdapter4.Fill(ds4);
                                stagedDataGridView.DataSource = ds4.Tables[0];
                                var maxLengthFieldArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

                                var selectMaxLengthColumnNumber = "USE " + databaseSelect.Text + " SELECT IMFF.FieldSeq FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null and ef.MaxLength is not null order by imff.FieldSeq";
                                var dataAdapter6 = new SqlDataAdapter(selectMaxLengthColumnNumber, conn);
                                var ds6 = new DataSet();
                                dataAdapter6.Fill(ds6);
                                stagedDataGridView.DataSource = ds6.Tables[0];
                                var maxLengthFieldColumnNumberArray = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

                                var selectMaxLengthValue = "USE " + databaseSelect.Text + " SELECT ef.maxlength FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null and ef.MaxLength is not null order by imff.FieldSeq";
                                var dataAdapter5 = new SqlDataAdapter(selectMaxLengthValue, conn);
                                var ds5 = new DataSet();
                                dataAdapter5.Fill(ds5);
                                stagedDataGridView.DataSource = ds5.Tables[0];
                                var maxLengthFieldArrayValue = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

                                var selectClientName = "USE " + databaseSelect.Text + " select optval from optset where OptName='ui.title.prefix'";
                                var dataAdapter7 = new SqlDataAdapter(selectClientName, conn);
                                var ds7 = new DataSet();
                                dataAdapter7.Fill(ds7);
                                stagedDataGridView.DataSource = ds7.Tables[0];
                                var clientName = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

                                var iffidArray = importformatDataGridView.Rows.Cast<DataGridViewRow>()
                                        .Select(x => x.Cells[5].Value.ToString().Trim()).ToArray();

                                var seqArray = importformatDataGridView.Rows.Cast<DataGridViewRow>()
                                    .Select(x => x.Cells[6].Value.ToString().Trim()).ToArray();



                                //var reqArray = importformatDataGridView.Rows.Cast<DataGridViewRow>()
                                //    .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();

                                int[] fieldsThatAreCodesArrayColumnCount = Array.ConvertAll(fieldsThatAreCodesArray, s => int.Parse(s));

                                ArrayList codeValueArray = new ArrayList();
                                //this foreach gets the values for all of the codes
                                foreach (var s in codeArray)
                                {
                                    var select2 = "USE " + databaseSelect.Text + "  select recval from codset where rectype="+"'"+s+"'";
                                    var dataAdapter2 = new SqlDataAdapter(select2, conn);
                                    var ds2 = new DataSet();
                                    dataAdapter2.Fill(ds2);
                                    stagedDataGridView.DataSource = ds2.Tables[0];

                                    foreach (DataGridViewRow dr in stagedDataGridView.Rows)
                                    {
                                        codeValueArray.Add(dr.Cells[0].Value);
                                    }
                                }

                                foreach (var value in clientName)
                                {
                                    tw.WriteLine("Client: " + value);
                                }
                                tw.WriteLine(".");
                                tw.WriteLine("---DATA THAT IS USED---");
                                foreach (DataGridViewRow dr in importformatDataGridView.Rows)
                                {
                                    bool checkBoxValue = Convert.ToBoolean(dr.Cells[0].Value);
                                    //tw.WriteLine("checkbox: " + checkBoxValue);
                                }
                                tw.WriteLine("---THIS IS DATA PULLED FROM "+databaseSelect.Text+"---");
                                foreach (var value in codeArray)
                                {
                                    tw.WriteLine("Code: " + value);
                                }
                                foreach (var value in codeValueArray)
                                {
                                    tw.WriteLine("Code Value: "+value);
                                }
                                foreach (int value in fieldsThatAreCodesArrayColumnCount)
                                {
                                    tw.WriteLine("Columns with Codes: " + value);
                                }

                                foreach (var value in maxLengthFieldColumnNumberArray)
                                {
                                    tw.WriteLine("Columns with length restrictions: " + value);
                                }

                                foreach (var value in maxLengthFieldArrayValue)
                                {
                                    tw.WriteLine("length restriction: " + value);
                                }



                                var intersect = fieldsThatAreCodesArray.Intersect(seqArray);
                                //int[] intIntersect = Array.ConvertAll(seqArray, s => int.Parse(s));
                                int[] intMaxLengthFieldArrayValue = Array.ConvertAll(maxLengthFieldArrayValue, s => int.Parse(s));



                                tw.WriteLine(".");
                                tw.WriteLine(".");
                                tw.WriteLine("ERROR LIST START");
                                tw.WriteLine("");



                                tw.WriteLine("Code Check");
                                int a = 0;
                                foreach (var s in iffidArray)
                                {
                                    a++;
                                        
                                    if (fieldsThatAreCodesArrayColumnCount.Contains(a) == true)
                                    {
                                        tw.WriteLine("\nCOLUMN " + a + ": " + s);//this is the header line in the output file
                                        for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)//this is the loop that spits out the errors
                                        {
                                            var value = importedfileDataGridView.Rows[i].Cells[a - 1].Value.ToString();
                                            if (codeValueArray.Contains(value) == false)
                                            {
                                                tw.WriteLine("Error at line " + (i + 1) + "." + " The value: '" + value + "' from your imported file does not exist in the database.");
                                            }
                                        }
                                    }
                                }

                                tw.WriteLine("Required Field Check");
                                int b = 0;
                                foreach (var s in iffidArray)
                                {
                                    a++;

                                    if (fieldsThatAreCodesArrayColumnCount.Contains(a) == true)
                                    {
                                        tw.WriteLine("\nCOLUMN " + a + ": " + s);//this is the header line in the output file
                                        for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)//this is the loop that spits out the errors
                                        {
                                            var value = importedfileDataGridView.Rows[i].Cells[a - 1].Value.ToString();
                                            if (codeValueArray.Contains(value) == false)
                                            {
                                                tw.WriteLine("Error at line " + (i + 1) + "." + " This column is required and you have a missing value.");
                                            }
                                        }
                                    }
                                }


                                tw.WriteLine("Max Length Check");
                                foreach (var s in seqArray)//cycle through every column
                                {
                                    if (maxLengthFieldColumnNumberArray.Contains(s) == true)//if one of the columns has a max length, enter this IF
                                    {
                                        
                                        int index = Array.IndexOf(seqArray, s);


                                        for (int j=0; j< importedfileDataGridView.Columns.Count; j++)
                                        {
                                            if (index==j)
                                            {
                                                b++;
                                                for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)//this is the loop that spits out the errors
                                                {
                                                    
                                                    var value = importedfileDataGridView.Rows[i].Cells[j].Value.ToString();
                                                    int valueLength = value.Length;
                                                    int maxValueLength = intMaxLengthFieldArrayValue[b-1];
                                                    if (valueLength > maxValueLength)
                                                    {
                                                        tw.WriteLine("Column: " + s);
                                                        tw.WriteLine("Error at line " + (i + 1) + "." + " The value: '" + value + "' from your imported file is "+valueLength+" characters long. This is too long.");
                                                    }
                                                }
                                            }
                                        }


                                    }
                                }





                                toolStripStatusLabel10.Text = importformatDataGridView.Rows.Count.ToString();
                                toolStripStatusLabel7.Text = stagedDataGridView.Rows.Count.ToString();
                                conn.Close();
                            }
                            catch { return; }

                            conn.Close();
                        }
                      
                        tw.WriteLine("EOF.");
                    }


                }
                MessageBox.Show("Import Format error file has been created. \nLocation: C:\\Program Files (x86)\\DataAnalysisTool\\Import Format Error Files", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + @">>>   Import Format error file has been created. Location: C:\Program Files (x86)\DataAnalysisTool\Import Format Error Files");
                progressBar1.MarqueeAnimationSpeed = 0;
                progressBar1.Refresh();
            }
        }
    }
}
