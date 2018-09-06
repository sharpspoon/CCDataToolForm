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

namespace DataAnalysisTool
{
    public partial class DataAnalysisTool
    {
        Importformat imp = new Importformat();
        //------------------SQL LOADER START------------------------------------------------------

        private void serverSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Loading load = new Loading();
            //load.ShowDialog();
            progressBar2.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            progressBar2.Value = 20;
            progressBar2.Value = 40;
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
                connectionStatus.Visible = true;
                richTextBox1.Text=richTextBox1.Text.Insert(0,Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                progressBar2.Value = 0;
                return;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar2.Value = 100;
        }

        private void serverSelect2_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar2.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            progressBar2.Value = 20;
            progressBar2.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect2.Text + "; Initial Catalog = master; Integrated Security = True");
            try
            {
                conn.Open();
                SqlCommand sc = new SqlCommand("SELECT name FROM [master].[sys].[databases] where name <> 'master' and name <> 'tempdb' and name <> 'model' and name <> 'msdb' and name <> 'DBAtools'", conn);
                SqlDataReader reader;
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                databaseSelect2.DataSource = dt;
                databaseSelect2.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                progressBar2.Value = 0;
                return;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar2.Value = 100;
        }

        private void serverSelect3_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Loading load = new Loading();
            //load.ShowDialog();
            progressBar2.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            progressBar2.Value = 20;
            progressBar2.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect3.Text + "; Initial Catalog = master; Integrated Security = True");
            try
            {
                conn.Open();
                SqlCommand sc = new SqlCommand("SELECT name FROM [master].[sys].[databases] where name <> 'master' and name <> 'tempdb' and name <> 'model' and name <> 'msdb' and name <> 'DBAtools'", conn);
                SqlDataReader reader;
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                databaseSelect3.DataSource = dt;
                databaseSelect3.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect3.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                progressBar2.Value = 0;
                return;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar2.Value = 100;
        }

        private void serverSelect4_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Loading load = new Loading();
            //load.ShowDialog();
            progressBar2.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            progressBar2.Value = 20;
            progressBar2.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect4.Text + "; Initial Catalog = master; Integrated Security = True");
            try
            {
                conn.Open();
                SqlCommand sc = new SqlCommand("SELECT name FROM [master].[sys].[databases] where name <> 'master' and name <> 'tempdb' and name <> 'model' and name <> 'msdb' and name <> 'DBAtools'", conn);
                SqlDataReader reader;
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                databaseSelect4.DataSource = dt;
                databaseSelect4.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect4.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                progressBar2.Value = 0;
                return;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar2.Value = 100;
        }

        private void runquery_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            System.Threading.Thread.Sleep(25);
            progressBar2.Value = 20;
            progressBar2.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect2.Text + "; Initial Catalog = master; Integrated Security = True");
            
            try
            {
                string ID = databaseSelect2.SelectedValue.ToString();
                conn.Open();
                var select = "USE " + databaseSelect2.Text + " " + queryWindow.Text;
                if (queryWindow.Text.Equals("select * from tranhis", StringComparison.InvariantCultureIgnoreCase))
                {
                    DialogResult result = MessageBox.Show("Performing a SELECT * FROM TRANHIS is insane. Continue?", "Data Analysis Tool", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (result == DialogResult.No)
                    {
                        progressBar1.MarqueeAnimationSpeed = 0;
                        progressBar2.Value = 0;
                        return;
                    }
                }
                var conn2 = new SqlConnection(@"Data Source = " + serverSelect2.Text + "; Initial Catalog = master; Integrated Security = True");
                var dataAdapter = new SqlDataAdapter(select, conn2);
                var commandBuilder = new SqlCommandBuilder(dataAdapter);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                dataGridView2.ReadOnly = true;
                dataGridView2.DataSource = ds.Tables[0];
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                conn.Close();
                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Running query against: " + databaseSelect2.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to run query. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                progressBar2.Value = 0;
                return;
            }
            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar2.Value = 100;
        }

        private void databaseSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            progressBar2.Value = 20;
            progressBar2.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc = new SqlCommand("use " + databaseSelect.Text + " SELECT table_name AS name FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE='BASE TABLE' order by TABLE_NAME", conn);
            SqlCommand scVersion = new SqlCommand("use " + databaseSelect.Text + " SELECT codetype FROM entityfield", conn);
            SqlDataReader reader;

            try
            {
                reader = scVersion.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                icmVersion.Visible = true;
                icmVersion.Text = "v.7.0";
            }
            catch
            {
                icmVersion.Visible = true;
                icmVersion.Text = "v.2018";
            }

            try
            {
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                tableSelect.DataSource = dt;
                tableSelect.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading database: " + databaseSelect.Text + "...Done.");
                toolStripStatusLabel5.Visible = true;
                toolStripStatusLabel6.Visible = true;
                toolStripStatusLabel7.Visible = true;
            }
            catch
            {
                return;
            }

            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar2.Value = 100;
        }

        //databaseSelect2 not used right now

        private void databaseSelect3_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            progressBar2.Value = 20;
            progressBar2.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect3.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc = new SqlCommand("use " + databaseSelect3.Text + " SELECT reportid as name FROM jasperreport  order by name", conn);
            SqlCommand sc2 = new SqlCommand("use " + databaseSelect3.Text + " SELECT statementtemplateid AS name FROM statementtemplate order by name", conn);
            
            SqlDataReader reader;

            try
            {
                if (reportRadio.Checked == true)
                {
                    reader = sc.ExecuteReader();
                }
                else
                {
                    reader = sc2.ExecuteReader();
                }
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                reportStatementSelect.DataSource = dt;
                reportStatementSelect.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading database: " + databaseSelect.Text + "...Done.");
                toolStripStatusLabel5.Visible = true;
                toolStripStatusLabel6.Visible = true;
                toolStripStatusLabel7.Visible = true;
            }
            catch
            {
                return;
            }

            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar2.Value = 100;
        }

        private void databaseSelect4_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("should not hit this");
            //payoutSelect.SelectedIndex = -1;
            //payoutTypeSelect.SelectedIndex = -1;
            progressBar1.MarqueeAnimationSpeed = 1;
            progressBar2.Value = 20;
            progressBar2.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect4.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc = new SqlCommand("use " + databaseSelect4.Text + " SELECT payouttypeid as name FROM payouttype  order by name", conn);
            SqlDataReader reader;

            try
            {
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                payoutTypeSelect.DataSource = dt;
                payoutTypeSelect.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading database: " + databaseSelect.Text + "...Done.");
                toolStripStatusLabel5.Visible = true;
                toolStripStatusLabel6.Visible = true;
                toolStripStatusLabel7.Visible = true;
            }
            catch
            {
                conn.Close();
                progressBar1.MarqueeAnimationSpeed = 0;
                progressBar2.Value = 0;
                return;
            }
            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar2.Value = 100;
        }

        private void payoutTypeSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            progressBar2.Value = 20;
            progressBar2.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect4.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc1 = new SqlCommand("use " + databaseSelect4.Text + " select distinct datfrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and rl.finalizestatus='p' order by 1 desc", conn);
            SqlCommand sc2 = new SqlCommand("use " + databaseSelect4.Text + " select distinct datfrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and rl.finalizestatus='f' order by 1 desc", conn);
            SqlCommand sc3 = new SqlCommand("use " + databaseSelect4.Text + " select distinct datfrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and rl.finalizestatus='r' order by 1 desc", conn);

            SqlDataReader reader;

            try
            {
                if (pendingRadioButton.Checked == true)
                {
                    reader = sc1.ExecuteReader();
                }
                else if (finalizedRadioButton.Checked == true)
                {
                    reader = sc2.ExecuteReader();
                }
                else if (reversedRadioButton.Checked == true)
                {
                    reader = sc3.ExecuteReader();
                }
                else
                {
                    return;
                }
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                payoutSelect.DataSource = dt;
                payoutSelect.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading payouts: " + payoutTypeSelect.Text + "...Done.");
                toolStripStatusLabel5.Visible = true;
                toolStripStatusLabel6.Visible = true;
                toolStripStatusLabel7.Visible = true;
            }
            catch
            {
                conn.Close();
                progressBar1.MarqueeAnimationSpeed = 0;
                progressBar2.Value = 0;
                return;
            }
            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar2.Value = 100;
        }

        private void tableSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            progressBar2.Value = 20;
            //System.Threading.Thread.Sleep(25);
            progressBar2.Value = 40;
            string ID = databaseSelect.SelectedValue.ToString();
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc;
            if (checkBox4.Checked == true)
            {
                sc = new SqlCommand("use " + databaseSelect.Text + " select importformatid as name from ImportFormat", conn);
            }
            else
            {
                sc = new SqlCommand("use " + databaseSelect.Text + " select importformatid as name from ImportFormat where prosta=1", conn);
            }
            
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
            catch
            {
                return;
            }

            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar2.Value = 100;
        }



        private void ifSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            progressBar2.Value = 20;
            progressBar2.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            try
            {
                var select = "USE " + databaseSelect.Text + " SELECT IMF.ImportFormatId,IMF.Delimiter,IMF.HeaderRows,IMF.RecType,IMFE.InEntName,IMFF.ImportFormatFieldId,IMFF.FieldSeq,IMFF.FieldLength,IMFF.IgnoreField, ef.* FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null order by imff.FieldSeq";
                var select2 = "USE " + databaseSelect.Text + " SELECT IMFF.ImportFormatFieldId FROM ImportFormat IMF INNER JOIN ImportFormatEntity IMFE ON IMF.ImportFormatNo= IMFE.ImportFormatNo INNER JOIN ImportFormatField IMFF ON IMF.ImportFormatNo = IMFF.ImportFormatNo  left JOIN EntityField EF ON ef.entname=imfe.inentname and ef.fldname=IMFF.ImportFormatFieldId where imf.importformatid = " + @"'" + ifSelect.Text + @"'" + "  and IMF.QBQueryNo is null order by imff.FieldSeq";
                var conn2 = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
                var dataAdapter = new SqlDataAdapter(select, conn2);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                importformatDataGridView.DataSource = ds.Tables[0];

                var iffidArray2 = importformatDataGridView.Rows.Cast<DataGridViewRow>()
                    .Select(x => x.Cells[5].Value.ToString().Trim()).ToArray();

                //gives me the InEntName of the import format
                var selectInEntName = "USE " + databaseSelect.Text + " select top 1 ife.InEntName from ImportFormat i inner join importformatentity ife on i.ImportFormatNo=ife.ImportFormatNo left join ImportFormatFieldMapping iffm on iffm.ImportFormatEntityNo=ife.ImportFormatEntityNo where i.ImportFormatId=" + @"'" + ifSelect.Text + @"'";
                var dataAdapter8 = new SqlDataAdapter(selectInEntName, conn);
                var ds8 = new DataSet();
                dataAdapter8.Fill(ds8);
                stagedDataGridView.DataSource = ds8.Tables[0];
                var inEntName = stagedDataGridView.Rows.Cast<DataGridViewRow>()
                        .Select(x => x.Cells[0].Value.ToString().Trim()).ToArray();
                var inEntNameVar  = stagedDataGridView.Rows[0].Cells[0].Value.ToString();

                if (inEntNameVar == "InAddress")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InAdjustmentHis")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InAssignment")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InBroker")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InBrokerAdj")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InBrokerContract")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InBrokerCustomer")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InBrokerDetail")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InBrokerHierarchy")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InBrokerHold")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InBrokerLicense")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InBrokerReserveHis")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InBrokerRoleBroker")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InBrokerVendor")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InCarrier")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InCertificate")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InCertificateDet")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InCmsMarx")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InCmsMmr")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InCmsTrr")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InCodSet")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InCustomer")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InCustomerApplication")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InCustomerMatch")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InCustPolicy")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InEducation")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InEntityRef")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InExtCrossRef")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InFile")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InFileImportFile")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InFileImportParm")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InFileImportRequest")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InFileRunList")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InIdentSet")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InMatchRule")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InPerfHis")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InPrepayBalanceAdjustment")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InProAppointment")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InProAppointmentDet")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InProBackground")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InProContract")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InProContractDet")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InProducer")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InProducts")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InProductsLicense")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InProInsurance")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InProLicense")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InProLicenseDet")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InterestDetail")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InterestSet")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InTimeSheet")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InTranDefault")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InTranHead")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InVendor")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }
                if (inEntNameVar == "InVoucher")
                {
                    label9.Text = inEntNameVar;
                    checkBox3.Visible = true;
                }

                reqListBox.Items.Clear();
                int a = 0;
                for (int i = 0; i < iffidArray2.Length; i++)
                {
                    a++;
                    reqListBox.Items.Add(a + ". " + iffidArray2[i].ToString());
                }

                dateListBox.Items.Clear();

                a = 0;
                for (int i = 0; i < iffidArray2.Length; i++)
                {
                    a++;
                    dateListBox.Items.Add(a+". "+iffidArray2[i].ToString());
                }
                conn.Close();

                toolStripStatusLabel7.Text = dataGridView2.Rows.Count.ToString();
                toolStripStatusLabel10.Text = importformatDataGridView.Rows.Count.ToString();
                richTextBox1.Text = richTextBox1.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading import format: " + ifSelect.Text + "...Done.");
                toolStripStatusLabel8.Visible = true;
                toolStripStatusLabel9.Visible = true;
                toolStripStatusLabel10.Visible = true;
            }
            catch
            {
                return;
            }

            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            progressBar2.Value = 100;
        }

        //------------------SQL LOADER END------------------------------------------------------
    }
}