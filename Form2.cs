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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            System.IO.Directory.CreateDirectory("C:\\Program Files (x86)\\CCDataTool\\ACTEKSOFT");
            string path = @"C:\\Program Files (x86)\\CCDataTool\\ACTEKSOFT\\LaunchWithACTEKSOFT.cmd";
            using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
            {
                using (TextWriter tw = new StreamWriter(fs))
                {
                    tw.WriteLine("taskkill /IM CCDataImportTool.exe /F");
                    tw.WriteLine("C:\\Windows\\System32\\runas.exe /user:ACTEKSOFT\\"+textBox7.Text+" /netonly CCDataImportTool.exe");
                }
                System.Diagnostics.Process.Start("C:\\Program Files (x86)\\CCDataTool\\ACTEKSOFT\\LaunchWithACTEKSOFT.cmd");

            }
        }
    }
}
