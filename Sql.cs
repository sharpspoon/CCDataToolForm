﻿using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.IO.Compression;

namespace DataAnalysisTool
{
    public partial class DataAnalysisTool
    {
        Importformat imp = new Importformat();



        //------------------SQL LOADER START------------------------------------------------------

        private void serverSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
            
            try
            {
                conn.Open();
                SqlCommand sc = new SqlCommand("SELECT name FROM [master].[sys].[databases] where name <> 'master' and name <> 'tempdb' and name <> 'model' and name <> 'msdb' and name <> 'DBAtools'", conn);
                SqlDataReader reader;
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                databaseSelect.DataSource = dt;
                databaseSelect.DisplayMember = "name";
                conn.Close();
                richTextBox1.Text=richTextBox1.Text.Insert(0,Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                return;
            }
        }

        private void databaseSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
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
                tableSelect.DataSource = dt;
                tableSelect.DisplayMember = "name";
                conn.Close();
                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading database: " + databaseSelect.Text + "...Done.");
                toolStripStatusLabel5.Visible = true;
                toolStripStatusLabel6.Visible = true;
                toolStripStatusLabel7.Visible = true;
            }
            catch { return; }

            conn.Close();
        }

        private void tableSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            string ID = databaseSelect.SelectedValue.ToString();
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc = new SqlCommand("use " + databaseSelect.Text + " select importformatid as name from ImportFormat", conn);
            SqlDataReader reader;
            try
            {
                var select = "USE " + databaseSelect.Text + " SELECT top 20000 * FROM " + tableSelect.Text;
                var conn2 = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
                var dataAdapter = new SqlDataAdapter(select, conn2);
                var commandBuilder = new SqlCommandBuilder(dataAdapter);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                dataGridView2.ReadOnly = true;
                dataGridView2.DataSource = ds.Tables[0];
                toolStripStatusLabel7.Text = dataGridView2.Rows.Count.ToString();

                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                ifSelect.DataSource = dt;
                ifSelect.DisplayMember = "name";
                conn.Close();
                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading table: " + tableSelect.Text + "...Done.");
            }
            catch { return; }

            conn.Close();
        }

        private void ifSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc = new SqlCommand("use " + databaseSelect.Text + " select importformatid as name from ImportFormat", conn);
            SqlDataReader reader;
            try
            {

                var select = "USE " + databaseSelect.Text + " SELECT IMF.ImportFormatId,IMF.Delimiter,IMF.HeaderRows,IMF.RecType,IMFE.InEntName,IMFF.ImportFormatFieldId,IMFF.FieldSeq,IMFF.FieldLength,IMFF.IgnoreField, ef.* FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null order by imff.FieldSeq";
                var conn2 = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
                var dataAdapter = new SqlDataAdapter(select, conn2);
                var commandBuilder = new SqlCommandBuilder(dataAdapter);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                dataGridView3.ReadOnly = true;
                dataGridView3.DataSource = ds.Tables[0];
                toolStripStatusLabel7.Text = dataGridView2.Rows.Count.ToString();
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                toolStripStatusLabel10.Text = dataGridView3.Rows.Count.ToString();
                conn.Close();
                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading import format: " + ifSelect.Text + "...Done.");
                toolStripStatusLabel8.Visible = true;
                toolStripStatusLabel9.Visible = true;
                toolStripStatusLabel10.Visible = true;
            }
            catch { return; }

            conn.Close();
        }

        //------------------SQL LOADER END------------------------------------------------------

    }
}