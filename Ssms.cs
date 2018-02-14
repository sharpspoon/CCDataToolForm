using System;
using System.Globalization;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.IO;
using System.DirectoryServices.ActiveDirectory;
using System.Web;

namespace CCDataImportTool
{
    public partial class Ssms : Form
    {
        public Ssms()
        {
            InitializeComponent();
        }

        private void Ssms_Load(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            System.IO.Directory.CreateDirectory("C:\\Program Files (x86)\\CCDataTool\\SSMS");
            string path = @"C:\\Program Files (x86)\\CCDataTool\\SSMS\\LaunchWithACTEKSOFT.cmd";
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (TextWriter tw = new StreamWriter(fs))
                {
                    tw.WriteLine("C:\\Windows\\System32\\runas.exe /user:ACTEKSOFT\\" + textBox7.Text + " /netonly "+ @"""C:\Program Files (x86)\Microsoft SQL Server\110\Tools\Binn\ManagementStudio\Ssms.exe""");
                }
                System.Diagnostics.Process.Start("C:\\Program Files (x86)\\CCDataTool\\SSMS\\LaunchWithACTEKSOFT.cmd");
                this.Close();

            }
        }
    }
}
