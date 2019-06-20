using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SAPDataAnalysisTool
{
    public partial class Loading : Form
    {
        public Loading()
        {

            InitializeComponent();
            this.BackColor = Color.LimeGreen;
            this.TransparencyKey = Color.LimeGreen;



        }

        private void Loading_Load(object sender, EventArgs e)
        {

        }

        private void Loading_Leave(object sender, EventArgs e)
        {
            //this.Close();
        }

        private void Loading_VisibleChanged(object sender, EventArgs e)
        {
            //this.Close();
        }

        private void Loading_Deactivate(object sender, EventArgs e)
        {

        }

        private void Loading_Shown(object sender, EventArgs e)
        {
            this.Refresh();
            SqlConnection conn = new SqlConnection(@"Data Source = IcmImpDb1.cci.caldsaas.local\Imp1; Initial Catalog = master; Integrated Security = True");
            try
            {
                //MessageBox.Show("asdf");
                conn.Open();
                //SqlCommand sc = new SqlCommand("SELECT name FROM [master].[sys].[databases] where name <> 'master' and name <> 'tempdb' and name <> 'model' and name <> 'msdb' and name <> 'DBAtools'", conn);
                //SqlDataReader reader;
                //reader = sc.ExecuteReader();
                //DataTable dt = new DataTable();
                //dt.Columns.Add("name", typeof(string));
                //dt.Load(reader);
                conn.Close();
            }
            catch
            {
                MessageBox.Show("Unable to connect to the server. Ensure you are connected with ACTEK", "Data Analysis Tool", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

                conn.Close();
            }
            //System.Threading.Thread.Sleep(5000);


            this.Close();


        }

        private void Loading_BackColorChanged(object sender, EventArgs e)
        {
        }
    }
}
