using System;
using System.Globalization;
using System.Threading;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Diagnostics;
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
    public partial class acteksoft : Form
    {
        public acteksoft()
        {
            InitializeComponent();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            var processdir = Environment.CurrentDirectory; 
            //string fullPath = process.;
            System.IO.Directory.CreateDirectory("C:\\Program Files (x86)\\CCDataTool\\ACTEKSOFT");
            string path = @"C:\\Program Files (x86)\\CCDataTool\\ACTEKSOFT\\LaunchWithACTEKSOFT.cmd";
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (TextWriter tw = new StreamWriter(fs))
                {
                    tw.WriteLine(@"md C:\Program Files(x86)\CCDataTool\Data");
                    tw.WriteLine("robocopy "+processdir+" "+@"""C:\Program Files (x86)\CCDataTool\Data"""+@" /MIR");
                    tw.WriteLine("taskkill /IM CCDataImportTool.exe /F");
                    tw.WriteLine("C:\\Windows\\System32\\runas.exe /user:ACTEKSOFT\\"+textBox7.Text+ @" /netonly ""C:\Program Files (x86)\CCDataTool\Data\CCDataImportTool.exe""");  //Environment.UserName
                }
                System.Diagnostics.Process.Start("C:\\Program Files (x86)\\CCDataTool\\ACTEKSOFT\\LaunchWithACTEKSOFT.cmd");

            }
        }

        private void acteksoft_Load(object sender, EventArgs e)
        {

        }
    }
}
