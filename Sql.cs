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
using System.Threading;

namespace SAPDataAnalysisTool
{
    public partial class SAPDataAnalysisTool
    {
        Importformat imp = new Importformat();
        //------------------SQL LOADER START------------------------------------------------------

        private void serverSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            importFormatProgressBar.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            importFormatProgressBar.Value = 20;
            importFormatProgressBar.Value = 40;
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
                systemLogTextBox.Text=systemLogTextBox.Text.Insert(0,Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                importFormatProgressBar.Value = 0;
                return;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            importFormatProgressBar.Value = 100;
        }

        private void serverSelect2_SelectedIndexChanged(object sender, EventArgs e)
        {

            sqlQueryProgressBar.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            sqlQueryProgressBar.Value = 20;
            sqlQueryProgressBar.Value = 40;
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
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                sqlQueryProgressBar.Value = 0;
                return;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            sqlQueryProgressBar.Value = 100;
        }



        private void serverSelect4_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            importFormatProgressBar.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            benchmarkProgressBar.Value = 20;
            benchmarkProgressBar.Value = 40;
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
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect4.Text + "...Done.");
                benchmarkProgressBar.Value = 100;
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                benchmarkProgressBar.Value = 0;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
        }


        private void serverSelect5_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            apiReadinessProgressBar.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            apiReadinessProgressBar.Value = 20;
            apiReadinessProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect5.Text + "; Initial Catalog = master; Integrated Security = True");
            try
            {
                conn.Open();
                SqlCommand sc = new SqlCommand("SELECT name FROM [master].[sys].[databases] where name <> 'master' and name <> 'tempdb' and name <> 'model' and name <> 'msdb' and name <> 'DBAtools'", conn);
                SqlDataReader reader;
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                databaseSelect5.DataSource = dt;
                databaseSelect5.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect5.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                apiReadinessProgressBar.Value = 0;
                return;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            apiReadinessProgressBar.Value = 100;
        }
        private void serverSelect6_SelectedIndexChanged(object sender, EventArgs e)
        {
            envChangesProgressBar.Value = 0;
            progressBar1.MarqueeAnimationSpeed = 1;
            envChangesProgressBar.Value = 20;
            envChangesProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect6.Text + "; Initial Catalog = master; Integrated Security = True");
            try
            {
                conn.Open();
                SqlCommand sc = new SqlCommand("SELECT name FROM [master].[sys].[databases] where name <> 'master' and name <> 'tempdb' and name <> 'model' and name <> 'msdb' and name <> 'DBAtools'", conn);
                SqlDataReader reader;
                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                databaseSelect6.DataSource = dt;
                databaseSelect6.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading SQL server: " + serverSelect6.Text + "...Done.");
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                envChangesProgressBar.Value = 0;
                return;
            }
            progressBar1.MarqueeAnimationSpeed = 0;
            envChangesProgressBar.Value = 100;
        }

        private void runquery_Click(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            System.Threading.Thread.Sleep(25);
            sqlQueryProgressBar.Value = 20;
            sqlQueryProgressBar.Value = 40;
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
                        sqlQueryProgressBar.Value = 0;
                        return;
                    }
                }
                var conn2 = new SqlConnection(@"Data Source = " + serverSelect2.Text + "; Initial Catalog = master; Integrated Security = True");
                var dataAdapter = new SqlDataAdapter(select, conn2);
                var commandBuilder = new SqlCommandBuilder(dataAdapter);
                var ds = new DataSet();
                dataAdapter.Fill(ds);
                sqlQueryDataGridView.ReadOnly = true;
                sqlQueryDataGridView.DataSource = ds.Tables[0];
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                conn.Close();
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Running query against: " + databaseSelect2.Text + "...Done.");
                sqlQueryProgressBar.Value = 100;
            }
            catch
            {
                conn.Close();
                MessageBox.Show("Unable to run query. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                progressBar1.MarqueeAnimationSpeed = 0;
                sqlQueryProgressBar.Value = 0;
                return;
            }
            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            importFormatProgressBar.Value = 100;
        }

        private void databaseSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            importFormatProgressBar.Value = 20;
            importFormatProgressBar.Value = 40;
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
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading database: " + databaseSelect.Text + "...Done.");
                seperator3ToolStripStatusLabel.Visible = true;
                sqlRowCountToolStripStatusLabel.Visible = true;
                sqlCounterToolStripStatusLabel.Visible = true;
            }
            catch
            {
                return;
            }

            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            importFormatProgressBar.Value = 100;
        }

        //databaseSelect2 not used right now
        //databaseSelect3 not used right now

        private void databaseSelect4_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("should not hit this");
            //payoutSelect.SelectedIndex = -1;
            //payoutTypeSelect.SelectedIndex = -1;
            progressBar1.MarqueeAnimationSpeed = 1;
            benchmarkProgressBar.Value = 20;
            benchmarkProgressBar.Value = 40;
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
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading database: " + databaseSelect.Text + "...Done.");
                seperator3ToolStripStatusLabel.Visible = true;
                sqlRowCountToolStripStatusLabel.Visible = true;
                sqlCounterToolStripStatusLabel.Visible = true;
            }
            catch
            {
                conn.Close();
                progressBar1.MarqueeAnimationSpeed = 0;
                benchmarkProgressBar.Value = 0;
                return;
            }
            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            benchmarkProgressBar.Value = 100;
        }


        private void payoutTypeSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            benchmarkProgressBar.Value = 20;
            benchmarkProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect4.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            //data select
            SqlDataReader reader;
            SqlCommand sc1 = new SqlCommand("use " + databaseSelect4.Text + " select distinct datfrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and rl.finalizestatus='p' order by 1 desc", conn);
            SqlCommand sc2 = new SqlCommand("use " + databaseSelect4.Text + " select distinct datfrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and rl.finalizestatus='f' order by 1 desc", conn);
            SqlCommand sc3 = new SqlCommand("use " + databaseSelect4.Text + " select distinct datfrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and rl.finalizestatus='r' order by 1 desc", conn);
            

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
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading payouts: " + payoutTypeSelect.Text + "...Done.");
                seperator3ToolStripStatusLabel.Visible = true;
                sqlRowCountToolStripStatusLabel.Visible = true;
                sqlCounterToolStripStatusLabel.Visible = true;
            }
            catch
            {
                conn.Close();
                progressBar1.MarqueeAnimationSpeed = 0;
                benchmarkProgressBar.Value = 0;
                return;
            }
            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            benchmarkProgressBar.Value = 100;
        }

        private void payoutSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            benchmarkProgressBar.Value = 20;
            benchmarkProgressBar.Value = 40;
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect4.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            //data select
            SqlDataReader reader;
            SqlCommand sc1 = new SqlCommand("use " + databaseSelect4.Text + " select distinct timefrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and DatFrom='" + payoutSelect.Text + "' and rl.finalizestatus='p' order by 1 desc", conn);
            SqlCommand sc2 = new SqlCommand("use " + databaseSelect4.Text + " select distinct timefrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and DatFrom='" + payoutSelect.Text + "' and rl.finalizestatus='f' order by 1 desc", conn);
            SqlCommand sc3 = new SqlCommand("use " + databaseSelect4.Text + " select distinct timefrom as name from RunList rl inner join rundet rd on rd.runlistno=rl.runlistno where rd.ItemName='PayoutTypeNo' and rd.ItemValue=(select payouttypeno from PayoutType where payouttypeid='" + payoutTypeSelect.Text + "') and rl.rectype='pay' and DatFrom='" + payoutSelect.Text + "' and rl.finalizestatus='r' order by 1 desc", conn);


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
                payoutTimeSelect.DataSource = dt;
                payoutTimeSelect.DisplayMember = "name";
                conn.Close();
                connectionStatus.Visible = true;
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading payouts: " + payoutTypeSelect.Text + "...Done.");
                seperator3ToolStripStatusLabel.Visible = true;
                sqlRowCountToolStripStatusLabel.Visible = true;
                sqlCounterToolStripStatusLabel.Visible = true;
            }
            catch
            {
                conn.Close();
                progressBar1.MarqueeAnimationSpeed = 0;
                benchmarkProgressBar.Value = 0;
                return;
            }
            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            benchmarkProgressBar.Value = 100;
        }

        private void tableSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            importFormatProgressBar.Value = 20;
            importFormatProgressBar.Value = 40;
            string ID = databaseSelect.SelectedValue.ToString();
            SqlConnection conn = new SqlConnection(@"Data Source = " + serverSelect.Text + "; Initial Catalog = master; Integrated Security = True");
            conn.Open();
            SqlCommand sc;
            if (importFormatShowOpenImportFormatsButton.Checked == true)
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
                sqlCounterToolStripStatusLabel.Text = dataGridView2.Rows.Count.ToString();

                reader = sc.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Columns.Add("name", typeof(string));
                dt.Load(reader);
                ifSelect.DataSource = dt;
                ifSelect.DisplayMember = "name";
                conn.Close();
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading table: " + tableSelect.Text + "...Done.");
            }
            catch
            {
                return;
            }

            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            importFormatProgressBar.Value = 100;
        }



        private void ifSelect_SelectedIndexChanged(object sender, EventArgs e)
        {
            progressBar1.MarqueeAnimationSpeed = 1;
            importFormatProgressBar.Value = 20;
            importFormatProgressBar.Value = 40;
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
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InAdjustmentHis")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InAssignment")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBroker")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerAdj")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerContract")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerCustomer")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerDetail")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerHierarchy")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerHold")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerLicense")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerReserveHis")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerRoleBroker")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InBrokerVendor")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCarrier")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCertificate")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCertificateDet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCmsMarx")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCmsMmr")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCmsTrr")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCodSet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCustomer")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCustomerApplication")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCustomerMatch")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InCustPolicy")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InEducation")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InEntityRef")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InExtCrossRef")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InFile")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InFileImportFile")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InFileImportParm")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InFileImportRequest")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InFileRunList")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InIdentSet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InMatchRule")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InPerfHis")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InPrepayBalanceAdjustment")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProAppointment")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProAppointmentDet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProBackground")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProContract")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProContractDet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProducer")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProducts")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProductsLicense")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProInsurance")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProLicense")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InProLicenseDet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InterestDetail")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InterestSet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InTimeSheet")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InTranDefault")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InTranHead")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InVendor")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
                }
                if (inEntNameVar == "InVoucher")
                {
                    importFormatInTableLabel.Text = inEntNameVar;
                    importFormatDatabaseCheck1Button.Visible = true;
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

                sqlCounterToolStripStatusLabel.Text = dataGridView2.Rows.Count.ToString();
                importFormatRowCountToolStripStatusLabel.Text = importformatDataGridView.Rows.Count.ToString();
                systemLogTextBox.Text = systemLogTextBox.Text.Insert(0, Environment.NewLine + DateTime.Now + ">>>   Loading import format: " + ifSelect.Text + "...Done.");
                seperator2ToolStripStatusLabel.Visible = true;
                ifRowCountToolStripStatusLabel.Visible = true;
                importFormatRowCountToolStripStatusLabel.Visible = true;
            }
            catch
            {
                return;
            }

            conn.Close();
            progressBar1.MarqueeAnimationSpeed = 0;
            importFormatProgressBar.Value = 100;
        }

        //------------------SQL LOADER END------------------------------------------------------
    }
}