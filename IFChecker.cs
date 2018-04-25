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
using System.Diagnostics;
using System.Collections.Generic;

namespace DataAnalysisTool
{
    public partial class DataAnalysisTool
    {

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
                        tw.WriteLine("###########################################################################################");
                        tw.WriteLine("##################DataAnalysisTool - Beginning of Import Format Error File#################");
                        tw.WriteLine("###########################################################################################");
                        tw.WriteLine("Reading file...Done.");
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
                                Process.Start(path);
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
                                tw.WriteLine("");
                                tw.WriteLine("---DATA THAT IS USED---");
                                String reqItem;
                                foreach (Object selecteditem in reqListBox.SelectedItems)
                                {

                                    reqItem = selecteditem as String;
                                    int reqCurIndex = reqListBox.Items.IndexOf(reqItem);
                                    if (reqCurIndex >= 0)
                                    {
                                        tw.WriteLine("Required Column: " + reqItem);
                                    }
                                }

                                String dateItem;
                                foreach (Object selecteditem in dateListBox.SelectedItems)
                                {

                                    dateItem = selecteditem as String;
                                    int dateCurIndex = dateListBox.Items.IndexOf(dateItem);
                                    if (dateCurIndex >= 0)
                                    {
                                        tw.WriteLine("Date Column: " + dateItem);
                                    }
                                }
                                tw.WriteLine(dateFormat.Text);
                                tw.WriteLine("");

                                int a = 0;
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
                                int[] intMaxLengthFieldArrayValue = Array.ConvertAll(maxLengthFieldArrayValue, s => int.Parse(s));

                                tw.WriteLine("");
                                tw.WriteLine("---ERROR LIST START---");

                                tw.WriteLine("--Required Field Check--");

                                //String reqItem;
                                foreach (Object selecteditem in reqListBox.SelectedItems)
                                {
                                    reqItem = selecteditem as String;
                                    int reqCurIndex = reqListBox.Items.IndexOf(reqItem);
                                    if (reqCurIndex >= 0)
                                    {
                                        tw.WriteLine("Required Column: " + reqItem);

                                        for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                                        {
                                            try
                                            {
                                                var value = importedfileDataGridView.Rows[i].Cells[reqCurIndex].Value.ToString();
                                                if (string.IsNullOrWhiteSpace(value))
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "." + " This column is required and you have a missing value.");
                                                }
                                            }
                                            catch (Exception)
                                            {
                                                // If we have reached this far, then none of the cells were empty.
                                                tw.WriteLine("No NULL values found in column " + "'" + reqItem + "'");
                                            }
                                        }
                                    }
                                }
                                tw.WriteLine("");

                                tw.WriteLine("--Code Check--");
                                a = 0;
                                foreach (var s in iffidArray)
                                {
                                    a++;
                                        
                                    if (fieldsThatAreCodesArrayColumnCount.Contains(a) == true)
                                    {
                                        tw.WriteLine("\nCOLUMN " + a + ": " + s);//this is the header line in the output file
                                        for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)//this is the loop that spits out the errors
                                        {
                                            var value = importedfileDataGridView.Rows[i].Cells[a - 1].Value.ToString();
                                            if (codeValueArray.Contains(value) == false && value != "")
                                            {
                                                tw.WriteLine("Error at line " + (i + 1) + "." + " The value: '" + value + "' from your imported file does not exist in the database.");
                                            }
                                        }
                                    }
                                }
                                tw.WriteLine("");

                                tw.WriteLine("--Max Length Check--");
                                a = 0;
                                foreach (var s in seqArray)//cycle through every column
                                {
                                    if (maxLengthFieldColumnNumberArray.Contains(s) == true)//if one of the columns has a max length, enter this IF
                                    {
                                        int index = Array.IndexOf(seqArray, s);
                                        for (int j=0; j< importedfileDataGridView.Columns.Count; j++)
                                        {
                                            if (index==j)
                                            {
                                                a++;
                                                for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)//this is the loop that spits out the errors
                                                {
                                                    
                                                    var value = importedfileDataGridView.Rows[i].Cells[j].Value.ToString();
                                                    int valueLength = value.Length;
                                                    int maxValueLength = intMaxLengthFieldArrayValue[a-1];
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
                                tw.WriteLine("");
                                tw.WriteLine("--Date Format Check--");
                                
                                foreach (Object selecteditem in dateListBox.SelectedItems)
                                {
                                    dateItem = selecteditem as String;
                                    int dateCurIndex = dateListBox.Items.IndexOf(dateItem);
                                    if (dateComboBox1.Text == "" && dateComboBox2.Text=="" && dateComboBox3.Text=="")
                                    {
                                        MessageBox.Show("Your date format is NULL. Please create a date format using the dropdown menus.");
                                        return;
                                    }
                                    string dateFormat2 = dateFormat.Text.Remove(0, 13);



                                    int dateFormatLength = dateFormat2.Length;
                                    MessageBox.Show("dateFormat2=" + dateFormat2+ "dateFormatLength="+ dateFormatLength);
                                    if (dateCurIndex >= 0)
                                    {
                                        if (dateFormatLength == 0) {
                                            MessageBox.Show("Your date format cannot be empty if you are specifying a date column", "DataAnalysisTool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                                            return;
                                        }

                                        tw.WriteLine("Date Column: " + dateItem);

                                        for (int i = 0; i < importedfileDataGridView.Rows.Count; i++)
                                        {
                                            var value = importedfileDataGridView.Rows[i].Cells[dateCurIndex].Value.ToString();
                                            int valueLength = value.Length;
                                            if ((dateFormatLength) != valueLength)
                                            {
                                                
                                                tw.WriteLine("Error at line " + (i + 1) + "." + " Your date format does not match your specified "+dateFormat.Text+".");
                                            }

                                            if (dateFormat2 == "yyyymmdd")
                                            {
                                                int year = int.Parse(value.Substring(0, 4));
                                                int month = int.Parse(value.Substring(4, 2));
                                                int day = int.Parse(value.Substring(6, 2));

                                                if (year > 2200)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: "+dateFormat2);
                                                }

                                                if (month > 12)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (month < 01)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (day > 31)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (day < 01)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }
                                            }

                                            if (dateFormat2 == "yyyyddmm")
                                            {
                                                int year = int.Parse(value.Substring(0, 4));
                                                int month = int.Parse(value.Substring(6, 2));
                                                int day = int.Parse(value.Substring(4, 2));

                                                if (year > 2200)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (month > 12)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (month < 01)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (day > 31)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (day < 01)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }
                                            }

                                            if (dateFormat2 == "yyddmm")
                                            {
                                                int year = int.Parse(value.Substring(0, 2));
                                                int month = int.Parse(value.Substring(4, 2));
                                                int day = int.Parse(value.Substring(2, 2));

                                                if (year > 22)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (month > 12)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (month < 01)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (day > 31)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (day < 01)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }
                                            }

                                            if (dateFormat2 == "yymmdd")
                                            {
                                                int year = int.Parse(value.Substring(0, 2));
                                                int month = int.Parse(value.Substring(2, 2));
                                                int day = int.Parse(value.Substring(4, 2));

                                                if (year > 22)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (month > 12)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (month < 01)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (day > 31)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (day < 01)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }
                                            }

                                            if (dateFormat2 == "mmddyyyy")
                                            {
                                                int year = int.Parse(value.Substring(4, 4));
                                                int month = int.Parse(value.Substring(0, 2));
                                                int day = int.Parse(value.Substring(2, 2));

                                                if (year > 2200)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (month > 12)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (month < 01)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (day > 31)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (day < 01)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }
                                            }

                                            if (dateFormat2 == "mmyyyydd")
                                            {
                                                int year = int.Parse(value.Substring(2, 4));
                                                int month = int.Parse(value.Substring(0, 2));
                                                int day = int.Parse(value.Substring(6, 2));

                                                if (year > 2200)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The year is " + year + ", which is greater than 2200.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (month > 12)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is greater than 12.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (month < 01)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The month is " + month + ", which is less than 1.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (day > 31)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is greater than 31.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }

                                                if (day < 01)
                                                {
                                                    tw.WriteLine("Error at line " + (i + 1) + "\r\n" + "The day is " + day + ", which is less than 01.\r\nMake sure that the date is in the format: " + dateFormat2);
                                                }
                                            }
                                        }
                                    }
                                }
                                tw.WriteLine("");
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
                Process.Start(path);
            }
        }
    }
}
